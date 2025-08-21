<?php

/**
 * Replace %*varName*% placeholders in a DOCX, including images and tables, even when broken across runs,
 * without changing styles. Writes to a new output file.
 *
 * Requirements: ext-dom, ext-zip
 */
function replaceDocxPlaceholdersSmart(string $inputDocx, string $outputDocx, array $replacements): void
{
    if (!is_file($inputDocx)) {
        throw new RuntimeException("Input file not found: $inputDocx");
    }

    // Work on a copy so the template is untouched.
    if (!@copy($inputDocx, $outputDocx)) {
        throw new RuntimeException("Failed to copy template to: $outputDocx");
    }

    $zip = new ZipArchive();
    if ($zip->open($outputDocx) !== true) {
        throw new RuntimeException("Unable to open DOCX: $outputDocx");
    }

    // Load and update [Content_Types].xml
    $contentTypesPath = '[Content_Types].xml';
    $contentTypesXml = $zip->getFromName($contentTypesPath);
    if ($contentTypesXml === false) {
        $zip->close();
        throw new RuntimeException("Unable to load [Content_Types].xml");
    }
    $contentTypesDoc = new DOMDocument();
    $contentTypesDoc->preserveWhiteSpace = true;
    $contentTypesDoc->formatOutput = false;
    $contentTypesDoc->loadXML($contentTypesXml);
    $contentTypesXp = new DOMXPath($contentTypesDoc);
    $contentTypesXp->registerNamespace('ct', 'http://schemas.openxmlformats.org/package/2006/content-types');

    // Load and update document relationships
    $relsPath = 'word/_rels/document.xml.rels';
    $relsXml = $zip->getFromName($relsPath);
    if ($relsXml === false) {
        $zip->close();
        throw new RuntimeException("Unable to load relationships file: $relsPath");
    }
    $relsDoc = new DOMDocument();
    $relsDoc->preserveWhiteSpace = true;
    $relsDoc->formatOutput = false;
    $relsDoc->loadXML($relsXml);
    $relsXp = new DOMXPath($relsDoc);
    $relsXp->registerNamespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships');

    // Process the main document.xml plus any headers/footers
    $parts = ['word/document.xml'];
    for ($i = 1; $i <= 8; $i++) {
        foreach (["word/header$i.xml", "word/footer$i.xml"] as $p) {
            if ($zip->locateName($p) !== false) $parts[] = $p;
        }
    }

    $imageRelationships = []; // Track image IDs and paths
    $imageCounter = 1; // For generating unique rId and image names
    $imageExtensions = []; // Track image extensions for content types

    foreach ($parts as $partPath) {
        $index = $zip->locateName($partPath);
        if ($index === false) continue;

        $xmlData = $zip->getFromIndex($index);
        $newXml = replaceInWordXml($xmlData, $replacements, $zip, $relsDoc, $relsXp, $contentTypesDoc, $contentTypesXp, $imageRelationships, $imageCounter, $imageExtensions);

        // Overwrite the XML part
        $zip->deleteName($partPath);
        $zip->addFromString($partPath, $newXml);
    }

    // Update [Content_Types].xml with image types
    foreach ($imageExtensions as $ext => $_) {
        $mimeType = getImageMimeType($ext);
        if ($mimeType && !$contentTypesXp->query("//ct:Default[@Extension='$ext']")->length) {
            $defaultNode = $contentTypesDoc->createElementNS('http://schemas.openxmlformats.org/package/2006/content-types', 'ct:Default');
            $defaultNode->setAttribute('Extension', $ext);
            $defaultNode->setAttribute('ContentType', $mimeType);
            $contentTypesDoc->documentElement->appendChild($defaultNode);
        }
    }

    // Update relationships and content types XML
    $zip->deleteName($relsPath);
    $zip->addFromString($relsPath, $relsDoc->saveXML());
    $zip->deleteName($contentTypesPath);
    $zip->addFromString($contentTypesPath, $contentTypesDoc->saveXML());

    // Add images to the zip
    foreach ($imageRelationships as $rId => $imageInfo) {
        if (!file_exists($imageInfo['path'])) {
            $zip->close();
            throw new RuntimeException("Image file not found: {$imageInfo['path']}");
        }
        $zip->addFile($imageInfo['path'], "word/media/{$imageInfo['name']}");
    }

    $zip->close();
}

/**
 * Core logic: walk each paragraph, scan its w:t nodes in order,
 * find %*...*% (across nodes), and replace with text, image, or table structure.
 */
function replaceInWordXml(string $xmlData, array $replacements, ZipArchive $zip, DOMDocument $relsDoc, DOMXPath $relsXp, DOMDocument $contentTypesDoc, DOMXPath $contentTypesXp, array &$imageRelationships, int &$imageCounter, array &$imageExtensions): string
{
    $doc = new DOMDocument();
    $doc->preserveWhiteSpace = true;
    $doc->formatOutput = false;
    if (!$doc->loadXML($xmlData)) {
        throw new RuntimeException("Failed to load XML data");
    }

    $xp = new DOMXPath($doc);
    $xp->registerNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
    $xp->registerNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing');
    $xp->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
    $xp->registerNamespace('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture');
    $xp->registerNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

    $paragraphs = $xp->query('//w:p');
    foreach ($paragraphs as $p) {
        // Collect runs to avoid live NodeList issues
        $runNodes = $xp->query('.//w:r', $p);
        $runs = [];
        foreach ($runNodes as $run) {
            $tNode = $xp->query('w:t', $run)->item(0);
            $runs[] = ['run' => $run, 'textNode' => $tNode];
        }

        $i = 0;
        while ($i < count($runs)) {
            $startRun = $runs[$i]['run'];
            $startTextNode = $runs[$i]['textNode'];
            $startText = $startTextNode ? $startTextNode->nodeValue : '';
            $posStart = strpos($startText, '%*');

            if ($posStart === false) {
                $i++;
                continue;
            }

            // Found a start marker
            $startIndex = $i;
            $startOffset = $posStart;
            $varText = '';
            $endFound = false;
            $endIndex = $startIndex;
            $endOffset = null;

            // Check if end marker is in the same node
            $afterStart = substr($startText, $startOffset + 2);
            $posEndSame = strpos($afterStart, '*%');
            if ($posEndSame !== false) {
                // Same-node placeholder
                $varRaw = substr($afterStart, 0, $posEndSame);
                $trailing = substr($afterStart, $posEndSame + 2);
                $leading = substr($startText, 0, $startOffset);
                $varName = preg_replace('/\s+/', '', $varRaw);

                if (isset($replacements[$varName]) && is_array($replacements[$varName]) && isset($replacements[$varName]['image'])) {
                    // Image replacement
                    $imageData = $replacements[$varName];
                    $ext = strtolower(pathinfo($imageData['image'], PATHINFO_EXTENSION));
                    $imageExtensions[$ext] = true;
                    $rId = "rIdImg" . $imageCounter++;
                    $imageName = "image" . $imageCounter . "." . $ext;
                    $imageRelationships[$rId] = [
                        'path' => $imageData['image'],
                        'name' => $imageName
                    ];

                    // Add relationship
                    $relNode = $relsDoc->createElementNS('http://schemas.openxmlformats.org/package/2006/relationships', 'r:Relationship');
                    $relNode->setAttribute('Id', $rId);
                    $relNode->setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
                    $relNode->setAttribute('Target', "media/$imageName");
                    $relsDoc->documentElement->appendChild($relNode);

                    // Create image run
                    $imageXml = createImageXml($rId, $imageName, $imageData['width'] ?? 300, $imageData['height'] ?? 300);
                    $imageFragment = $doc->createDocumentFragment();
                    if (!$imageFragment->appendXML($imageXml)) {
                        throw new RuntimeException("Failed to append image XML for $varName");
                    }

                    // Replace the entire run
                    if ($startRun->parentNode) {
                        $startRun->parentNode->replaceChild($imageFragment, $startRun);
                    }
                } elseif (isset($replacements[$varName]) && is_array($replacements[$varName]) && is_array($replacements[$varName][0])) {
                    // Table replacement
                    $tableData = $replacements[$varName];
                    $tableXml = createTableXml($tableData);
                    $tableFragment = $doc->createDocumentFragment();
                    if (!$tableFragment->appendXML($tableXml)) {
                        throw new RuntimeException("Failed to append table XML for $varName");
                    }

                    // Replace the entire paragraph
                    if ($p->parentNode) {
                        $p->parentNode->replaceChild($tableFragment, $p);
                    }
                    break; // Exit loop since paragraph is replaced
                } else {
                    // Text replacement
                    $replacement = array_key_exists($varName, $replacements) ? (string)$replacements[$varName] : null;
                    if ($replacement !== null) {
                        $startTextNode->nodeValue = $leading . $replacement . $trailing;
                    } else {
                        $startTextNode->nodeValue = $leading . $varRaw . $trailing;
                    }
                }
                $i++;
                continue;
            }

            // Multi-node placeholder
            $varText .= $afterStart;
            for ($j = $startIndex + 1; $j < count($runs); $j++) {
                $endTextNode = $runs[$j]['textNode'];
                $endCandidate = $endTextNode ? $endTextNode->nodeValue : '';
                $posEnd = strpos($endCandidate, '*%');
                if ($posEnd === false) {
                    $varText .= $endCandidate;
                    continue;
                }
                $varText .= substr($endCandidate, 0, $posEnd);
                $endIndex = $j;
                $endOffset = $posEnd;
                $endFound = true;
                break;
            }

            if (!$endFound) {
                $i++;
                continue;
            }

            $leadingStart = substr($startText, 0, $startOffset);
            $endTextNode = $runs[$endIndex]['textNode'];
            $endNodeText = $endTextNode ? $endTextNode->nodeValue : '';
            $trailingEnd = substr($endNodeText, $endOffset + 2);
            $varName = preg_replace('/\s+/', '', $varText);

            if (isset($replacements[$varName]) && is_array($replacements[$varName]) && isset($replacements[$varName]['image'])) {
                // Image replacement
                $imageData = $replacements[$varName];
                $ext = strtolower(pathinfo($imageData['image'], PATHINFO_EXTENSION));
                $imageExtensions[$ext] = true;
                $rId = "rIdImg" . $imageCounter++;
                $imageName = "image" . $imageCounter . "." . $ext;
                $imageRelationships[$rId] = [
                    'path' => $imageData['image'],
                    'name' => $imageName
                ];

                // Add relationship
                $relNode = $relsDoc->createElementNS('http://schemas.openxmlformats.org/package/2006/relationships', 'r:Relationship');
                $relNode->setAttribute('Id', $rId);
                $relNode->setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
                $relNode->setAttribute('Target', "media/$imageName");
                $relsDoc->documentElement->appendChild($relNode);

                // Create image run
                $imageXml = createImageXml($rId, $imageName, $imageData['width'] ?? 300, $imageData['height'] ?? 300);
                $imageFragment = $doc->createDocumentFragment();
                if (!$imageFragment->appendXML($imageXml)) {
                    throw new RuntimeException("Failed to append image XML for $varName");
                }

                // Replace the start run with the image
                if ($startRun->parentNode) {
                    $startRun->parentNode->replaceChild($imageFragment, $startRun);
                }

                // Remove middle runs
                for ($k = $startIndex + 1; $k < $endIndex; $k++) {
                    if ($runs[$k]['run']->parentNode) {
                        $runs[$k]['run']->parentNode->removeChild($runs[$k]['run']);
                    }
                }

                // Update end run
                if ($endTextNode) {
                    $endTextNode->nodeValue = $trailingEnd;
                }
            } elseif (isset($replacements[$varName]) && is_array($replacements[$varName]) && is_array($replacements[$varName][0])) {
                // Table replacement
                $tableData = $replacements[$varName];
                $tableXml = createTableXml($tableData);
                $tableFragment = $doc->createDocumentFragment();
                if (!$tableFragment->appendXML($tableXml)) {
                    throw new RuntimeException("Failed to append table XML for $varName");
                }

                // Replace the entire paragraph
                if ($p->parentNode) {
                    $p->parentNode->replaceChild($tableFragment, $p);
                }
                break; // Exit loop since paragraph is replaced
            } else {
                // Text replacement
                $replacement = array_key_exists($varName, $replacements) ? (string)$replacements[$varName] : null;
                if ($replacement !== null) {
                    $startTextNode->nodeValue = $leadingStart . $replacement;
                } else {
                    $startTextNode->nodeValue = $leadingStart . $varText;
                }

                for ($k = $startIndex + 1; $k < $endIndex; $k++) {
                    if ($runs[$k]['textNode']) {
                        $runs[$k]['textNode']->nodeValue = '';
                    }
                }

                if ($endTextNode) {
                    $endTextNode->nodeValue = $trailingEnd;
                }
            }

            $i = $endIndex + 1;
        }
    }

    return $doc->saveXML();
}

/**
 * Generate WordprocessingML for an image with proper namespace declarations.
 */
function createImageXml(string $rId, string $imageName, int $widthPx, int $heightPx): string
{
    // Convert pixels to EMUs (1 px = 9525 EMUs)
    $widthEmu = $widthPx * 9525;
    $heightEmu = $heightPx * 9525;

    // Define all necessary namespaces in the root element
    return <<<XML
<w:r
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:drawing>
        <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="$widthEmu" cy="$heightEmu"/>
            <wp:docPr id="1" name="$imageName"/>
            <a:graphic>
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                    <pic:pic>
                        <pic:nvPicPr>
                            <pic:cNvPr id="0" name="$imageName"/>
                            <pic:cNvPicPr/>
                        </pic:nvPicPr>
                        <pic:blipFill>
                            <a:blip r:embed="$rId"/>
                            <a:stretch>
                                <a:fillRect/>
                            </a:stretch>
                        </pic:blipFill>
                        <pic:spPr>
                            <a:xfrm>
                                <a:off x="0" y="0"/>
                                <a:ext cx="$widthEmu" cy="$heightEmu"/>
                            </a:xfrm>
                            <a:prstGeom prst="rect">
                                <a:avLst/>
                            </a:prstGeom>
                        </pic:spPr>
                    </pic:pic>
                </a:graphicData>
            </a:graphic>
        </wp:inline>
    </w:drawing>
</w:r>
XML;
}

/**
 * Generate WordprocessingML for a table.
 */
function createTableXml(array $tableData): string
{
    if (empty($tableData)) {
        return '';
    }

    $xml = '<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
    $xml .= '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>';
    $xml .= '<w:tblGrid>';

    // Assume equal column widths for simplicity
    $colCount = count($tableData[0]);
    for ($i = 0; $i < $colCount; $i++) {
        $xml .= '<w:gridCol/>';
    }
    $xml .= '</w:tblGrid>';

    foreach ($tableData as $row) {
        $xml .= '<w:tr>';
        foreach ($row as $cell) {
            $xml .= '<w:tc>';
            $xml .= '<w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>';
            $xml .= '<w:p><w:r><w:t>' . htmlspecialchars((string)$cell, ENT_QUOTES, 'UTF-8') . '</w:t></w:r></w:p>';
            $xml .= '</w:tc>';
        }
        $xml .= '</w:tr>';
    }

    $xml .= '</w:tbl>';
    return $xml;
}

/**
 * Get MIME type for image extension.
 */
function getImageMimeType(string $extension): ?string
{
    $mimeTypes = [
        'jpg' => 'image/jpeg',
        'jpeg' => 'image/jpeg',
        'png' => 'image/png',
        'gif' => 'image/gif',
        'bmp' => 'image/bmp'
    ];
    return $mimeTypes[strtolower($extension)] ?? null;
}

// ------------------------------ Example usage ------------------------------
$values = [
    'variable' => 'Hello!....',
    'employerName' => 'Google Inc.',
    'CaseID' => 'BGV123456',
    'cName' => 'John Doe',
    'reportStatus' => 'Clear',
    'dob' => '1990-01-01',
    'requestDate' => '2025-08-20',
    'cID' => 'CID-98765',
    'reportDate' => '2025-08-21',
    'img1' => [
        'image' => 'C:\xampp\htdocs\test\pdfgenback\openxml\image1.jpg',
        'width' => 300,
        'height' => 300
    ],
    'img2' => [
        'image' => 'C:\xampp\htdocs\test\pdfgenback\openxml\image2.jpg',
        'width' => 600,
        'height' => 400
    ],
    'tableData' => [
        ['Column 1', 'Column 2', 'Column 3', 'Column 4'],
        ['Row 1 Col 1', 'Row 1 Col 2', 'Row 1 Col 3', 'Row 1 Col 4'],
        ['Row 2 Col 1', 'Row 2 Col 2', 'Row 2 Col 3', 'Row 2 Col 4']
    ],
];
$template = 'input.docx';
$output = 'output.docx';

replaceDocxPlaceholdersSmart($template, $output, $values);

echo "✅ Created $output\n";
?>