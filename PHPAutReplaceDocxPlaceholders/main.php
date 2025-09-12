<?php

/**
 * Replace %*varName*% placeholders in a DOCX, including images and tables, even when broken across runs,
 * and handle table row placeholders (e.g., arg_tX, arg_taX), adding new rows with preserved styling.
 * Supports optional cell background and font color for table replacements.
 * Writes to a new output file.
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
    $imageCounters = []; // Track counters for each image placeholder type

    foreach ($parts as $partPath) {
        $index = $zip->locateName($partPath);
        if ($index === false) continue;

        $xmlData = $zip->getFromIndex($index);
        $newXml = replaceInWordXml($xmlData, $replacements, $zip, $relsDoc, $relsXp, $contentTypesDoc, $contentTypesXp, $imageRelationships, $imageCounter, $imageExtensions, $imageCounters);

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
 * Core logic: walk each paragraph and table, scan w:t nodes for %*...*% placeholders,
 * replace with text, image, or table structure, and handle table row placeholders.
 */
function replaceInWordXml(string $xmlData, array $replacements, ZipArchive $zip, DOMDocument $relsDoc, DOMXPath $relsXp, DOMDocument $contentTypesDoc, DOMXPath $contentTypesXp, array &$imageRelationships, int &$imageCounter, array &$imageExtensions, array &$imageCounters): string
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

    // Process paragraphs for text, image, or full table replacements
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

                // Check for image placeholder with (i) pattern
                if (preg_match('/^(.+)\(i\)$/', $varName, $matches)) {
                    $baseImageName = $matches[1];
                    if (isset($replacements[$varName]) && is_array($replacements[$varName])) {
                        // Create multiple paragraphs with images
                        $imageXmlParts = [];
                        foreach ($replacements[$varName] as $imageKey => $imageData) {
                            if (is_array($imageData) && isset($imageData['image'])) {
                                $ext = strtolower(pathinfo($imageData['image'], PATHINFO_EXTENSION));
                                $imageExtensions[$ext] = true;
                                $rId = "rIdImg" . $imageCounter++;
                                
                                // Generate sequential filename
                                if (!isset($imageCounters[$baseImageName])) {
                                    $imageCounters[$baseImageName] = 1;
                                }
                                $imageName = $baseImageName . "-" . $imageCounters[$baseImageName]++ . "." . $ext;
                                
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
                                
                                // Create image paragraph
                                $imageXml = createImageXml($rId, $imageName, $imageData['width'] ?? 300, $imageData['height'] ?? 300);
                                $imageXmlParts[] = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' . $imageXml . '</w:p>';
                            }
                        }
                        
                        if (!empty($imageXmlParts)) {
                            $combinedXml = implode('', $imageXmlParts);
                            $imageFragment = $doc->createDocumentFragment();
                            if (!$imageFragment->appendXML($combinedXml)) {
                                throw new RuntimeException("Failed to append image XML for $varName");
                            }
                            
                            // Replace the entire paragraph
                            if ($p->parentNode) {
                                $p->parentNode->replaceChild($imageFragment, $p);
                            }
                            break; // Exit loop since paragraph is replaced
                        }
                    }
                } elseif (isset($replacements[$varName]) && is_array($replacements[$varName]) && isset($replacements[$varName]['image'])) {
                    // Single image replacement
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
                    // Full table replacement
                    $tableData = $replacements[$varName];
                    $tableXml = createTableXml($tableData, true); // Use template styling
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

            // Check for image placeholder with (i) pattern
            if (preg_match('/^(.+)\(i\)$/', $varName, $matches)) {
                $baseImageName = $matches[1];
                if (isset($replacements[$varName]) && is_array($replacements[$varName])) {
                    // Create multiple paragraphs with images
                    $imageXmlParts = [];
                    foreach ($replacements[$varName] as $imageKey => $imageData) {
                        if (is_array($imageData) && isset($imageData['image'])) {
                            $ext = strtolower(pathinfo($imageData['image'], PATHINFO_EXTENSION));
                            $imageExtensions[$ext] = true;
                            $rId = "rIdImg" . $imageCounter++;
                            
                            // Generate sequential filename
                            if (!isset($imageCounters[$baseImageName])) {
                                $imageCounters[$baseImageName] = 1;
                            }
                            $imageName = $baseImageName . "-" . $imageCounters[$baseImageName]++ . "." . $ext;
                            
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
                            
                            // Create image paragraph
                            $imageXml = createImageXml($rId, $imageName, $imageData['width'] ?? 300, $imageData['height'] ?? 300);
                            $imageXmlParts[] = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' . $imageXml . '</w:p>';
                        }
                    }
                    
                    if (!empty($imageXmlParts)) {
                        $combinedXml = implode('', $imageXmlParts);
                        $imageFragment = $doc->createDocumentFragment();
                        if (!$imageFragment->appendXML($combinedXml)) {
                            throw new RuntimeException("Failed to append image XML for $varName");
                        }
                        
                        // Replace the entire paragraph
                        if ($p->parentNode) {
                            $p->parentNode->replaceChild($imageFragment, $p);
                        }
                        break; // Exit loop since paragraph is replaced
                    }
                }
            } elseif (isset($replacements[$varName]) && is_array($replacements[$varName]) && isset($replacements[$varName]['image'])) {
                // Single image replacement
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
                // Full table replacement
                $tableData = $replacements[$varName];
                $tableXml = createTableXml($tableData, true); // Use template styling
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

    // Process tables for row-specific placeholder replacements
    $tables = $xp->query('//w:tbl');
    foreach ($tables as $table) {
        processTableRows($table, $replacements, $xp, $doc);
    }

    return $doc->saveXML();
}

/**
 * Process table rows to replace placeholders and add new rows with preserved or user-specified styling.
 */
function processTableRows(DOMElement $table, array $replacements, DOMXPath $xp, DOMDocument $doc): void
{
    $rows = $xp->query('w:tr', $table);
    $rowArray = [];
    foreach ($rows as $row) {
        $rowArray[] = $row;
    }

    // Get table properties for styling
    $tblPr = $xp->query('w:tblPr', $table)->item(0);
    $tblPrXml = $tblPr ? $doc->saveXML($tblPr) : '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>';

    foreach ($rowArray as $rowIndex => $row) {
        $cells = $xp->query('w:tc', $row);
        $hasPlaceholders = false;
        $placeholders = [];
        $cellProperties = [];
        $paraProperties = [];
        $runProperties = [];

        // Collect cell, paragraph, and run properties for styling
        foreach ($cells as $cellIndex => $cell) {
            $tcPr = $xp->query('w:tcPr', $cell)->item(0);
            $cellProperties[$cellIndex] = $tcPr ? $doc->saveXML($tcPr) : '<w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>';
            $para = $xp->query('w:p', $cell)->item(0);
            $pPr = $para ? $xp->query('w:pPr', $para)->item(0) : null;
            $paraProperties[$cellIndex] = $pPr ? $doc->saveXML($pPr) : '';
            $run = $xp->query('w:r', $para)->item(0);
            $rPr = $run ? $xp->query('w:rPr', $run)->item(0) : null;
            $runProperties[$cellIndex] = $rPr ? $doc->saveXML($rPr) : '';
        }

        // Check each cell for placeholders
        foreach ($cells as $cellIndex => $cell) {
            $textNodes = $xp->query('.//w:t', $cell);
            $cellText = '';
            foreach ($textNodes as $tNode) {
                $cellText .= $tNode->nodeValue;
            }
            preg_match_all('/%*\w+%*/', $cellText, $matches);
            if (!empty($matches[0])) {
                $hasPlaceholders = true;
                $placeholders[$cellIndex] = $matches[0];
            }
        }

        if ($hasPlaceholders) {
            // Identify the placeholder key (e.g., 'arg_t1' or 'arg_ta1')
            $firstPlaceholder = reset($placeholders)[0] ?? '';
            $varName = preg_replace('/%*|\*/', '', $firstPlaceholder);
            // Extract the base key (e.g., 'arg_t(i)' or 'arg_ta(i)')
            if (preg_match('/^(arg_t|arg_ta)\d+$/', $varName, $matches)) {
                $baseKey = $matches[1] . '(i)';
                if (isset($replacements[$baseKey]) && is_array($replacements[$baseKey]) && is_array($replacements[$baseKey][0])) {
                    $tableData = $replacements[$baseKey];

                    // Replace placeholders in the current row with the first row of data
                    if (!empty($tableData)) {
                        $firstRowData = array_shift($tableData);
                        foreach ($cells as $cellIndex => $cell) {
                            if (isset($placeholders[$cellIndex])) {
                                foreach ($placeholders[$cellIndex] as $placeholder) {
                                    $colVarName = preg_replace('/%*|\*/', '', $placeholder);
                                    $colIndex = (int)preg_replace('/^arg_t(a)?/', '', $colVarName) - 1;
                                    $cellValue = isset($firstRowData[$colIndex]) ? (string)(is_array($firstRowData[$colIndex]) ? ($firstRowData[$colIndex]['value'] ?? $firstRowData[$colIndex]) : $firstRowData[$colIndex]) : '';
                                    $textNodes = $xp->query('.//w:t', $cell);
                                    foreach ($textNodes as $tNode) {
                                        $tNode->nodeValue = str_replace($placeholder, $cellValue, $tNode->nodeValue);
                                    }
                                }
                            }
                        }
                    }

                    // Add new rows with preserved styling from the row above or user-specified colors
                    foreach ($tableData as $rowData) {
                        $newRowXml = '<w:tr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
                        foreach ($rowData as $cellIndex => $cellData) {
                            $cellValue = is_array($cellData) ? ($cellData['value'] ?? $cellData) : $cellData;
                            $bgColor = is_array($cellData) && isset($cellData['bgColor']) ? $cellData['bgColor'] : null;
                            $fontColor = is_array($cellData) && isset($cellData['fontColor']) ? $cellData['fontColor'] : null;

                            $newRowXml .= '<w:tc>';
                            $newRowXml .= $cellProperties[$cellIndex] ?? '<w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>';
                            if ($bgColor) {
                                $newRowXml .= "<w:tcPr><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"$bgColor\"/></w:tcPr>";
                            }
                            $newRowXml .= '<w:p>';
                            if (!empty($paraProperties[$cellIndex])) {
                                $newRowXml .= $paraProperties[$cellIndex];
                            }
                            $newRowXml .= '<w:r>';
                            // Preserve all run properties, only override color if specified
                            $runXml = $runProperties[$cellIndex] ?? '';
                            if ($fontColor && $runXml) {
                                // Modify existing run properties to update color only
                                $runDoc = new DOMDocument();
                                $runDoc->loadXML('<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' . $runXml . '</root>');
                                $runXp = new DOMXPath($runDoc);
                                $runXp->registerNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
                                $colorNode = $runXp->query('//w:color')->item(0);
                                if ($colorNode) {
                                    $colorNode->setAttribute('w:val', $fontColor);
                                } else {
                                    $rPrNode = $runXp->query('//w:rPr')->item(0);
                                    if ($rPrNode) {
                                        $newColorNode = $runDoc->createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:color');
                                        $newColorNode->setAttribute('w:val', $fontColor);
                                        $rPrNode->appendChild($newColorNode);
                                    }
                                }
                                $runXml = $runDoc->saveXML($runDoc->documentElement->firstChild);
                            } elseif ($fontColor) {
                                // If no run properties exist, create minimal with font color
                                $runXml = "<w:rPr><w:color w:val=\"$fontColor\"/><w:sz w:val=\"22\"/></w:rPr>";
                            }
                            $newRowXml .= $runXml;
                            $newRowXml .= '<w:t>' . htmlspecialchars((string)$cellValue, ENT_QUOTES, 'UTF-8') . '</w:t></w:r></w:p>';
                            $newRowXml .= '</w:tc>';
                        }
                        $newRowXml .= '</w:tr>';
                        $newRowFragment = $doc->createDocumentFragment();
                        if (!$newRowFragment->appendXML($newRowXml)) {
                            throw new RuntimeException("Failed to append new row XML");
                        }
                        $table->appendChild($newRowFragment);
                    }
                }
            }
        }
    }
}

/**
 * Generate WordprocessingML for a full table, optionally using template styling and user-specified colors.
 */
function createTableXml(array $tableData, bool $useTemplateStyle = false): string
{
    if (empty($tableData)) {
        return '';
    }

    $xml = '<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
    // Use template styling if available, otherwise default
    $xml .= $useTemplateStyle
        ? '<w:tblPr>
            <w:tblStyle w:val="TableGrid"/>
            <w:tblW w:w="5000" w:type="dxa"/>
            <w:tblBorders>
                <w:top w:val="single" w:color="000000" w:sz="12"/>
                <w:left w:val="single" w:color="000000" w:sz="12"/>
                <w:bottom w:val="single" w:color="000000" w:sz="12"/>
                <w:right w:val="single" w:color="000000" w:sz="12"/>
                <w:insideH w:val="single" w:color="000000" w:sz="12"/>
                <w:insideV w:val="single" w:color="000000" w:sz="12"/>
            </w:tblBorders>
          </w:tblPr>'
        : '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>';
    $xml .= '<w:tblGrid>';

    // Assume equal column widths for simplicity
    $colCount = count(isset($tableData[0]) && is_array($tableData[0]) ? $tableData[0] : []);
    for ($i = 0; $i < $colCount; $i++) {
        $xml .= '<w:gridCol w:w="' . ($useTemplateStyle ? 1250 : 0) . '"/>';
    }
    $xml .= '</w:tblGrid>';

    foreach ($tableData as $rowIndex => $row) {
        $xml .= '<w:tr>';
        foreach ($row as $cellIndex => $cellData) {
            $cellValue = is_array($cellData) ? ($cellData['value'] ?? $cellData) : $cellData;
            $bgColor = is_array($cellData) && isset($cellData['bgColor']) ? $cellData['bgColor'] : ($rowIndex == 0 && $useTemplateStyle ? '0070C0' : null);
            $fontColor = is_array($cellData) && isset($cellData['fontColor']) ? $cellData['fontColor'] : ($rowIndex == 0 && $useTemplateStyle ? 'FFFFFF' : ($rowIndex > 0 && $cellIndex == $colCount - 1 && $useTemplateStyle ? '00B050' : null));

            $xml .= '<w:tc>';
            $xml .= '<w:tcPr><w:tcW w:w="' . ($useTemplateStyle ? 1250 : 0) . '" w:type="dxa"/>';
            if ($bgColor) {
                $xml .= "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"$bgColor\"/>";
            }
            $xml .= '</w:tcPr>';
            $xml .= '<w:p><w:pPr><w:jc w:val="' . ($useTemplateStyle ? 'center' : 'left') . '"/></w:pPr><w:r>';
            // Preserve existing run properties, only override color if specified
            if ($fontColor) {
                $xml .= "<w:rPr><w:sz w:val=\"" . ($rowIndex == 0 && $useTemplateStyle ? 24 : 22) . "\"/><w:color w:val=\"$fontColor\"/></w:rPr>";
            } elseif ($useTemplateStyle) {
                $xml .= '<w:rPr><w:sz w:val="' . ($rowIndex == 0 ? 24 : 22) . '"/></w:rPr>';
            }
            $xml .= '<w:t>' . htmlspecialchars((string)$cellValue, ENT_QUOTES, 'UTF-8') . '</w:t>';
            $xml .= '</w:r></w:p>';
            $xml .= '</w:tc>';
        }
        $xml .= '</w:tr>';
    }

    $xml .= '</w:tbl>';
    return $xml;
}

/**
 * Generate WordprocessingML for a full table, optionally using template styling and user-specified colors.
 */
function createTableXml_old(array $tableData, bool $useTemplateStyle = false): string
{
    if (empty($tableData)) {
        return '';
    }

    $xml = '<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
    // Use template styling if available, otherwise default
    $xml .= $useTemplateStyle
        ? '<w:tblPr>
            <w:tblStyle w:val="TableGrid"/>
            <w:tblW w:w="5000" w:type="dxa"/>
            <w:tblBorders>
                <w:top w:val="single" w:color="000000" w:sz="12"/>
                <w:left w:val="single" w:color="000000" w:sz="12"/>
                <w:bottom w:val="single" w:color="000000" w:sz="12"/>
                <w:right w:val="single" w:color="000000" w:sz="12"/>
                <w:insideH w:val="single" w:color="000000" w:sz="12"/>
                <w:insideV w:val="single" w:color="000000" w:sz="12"/>
            </w:tblBorders>
          </w:tblPr>'
        : '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>';
    $xml .= '<w:tblGrid>';

    // Assume equal column widths for simplicity
    $colCount = count(isset($tableData[0]) ? $tableData[0] : []);
    for ($i = 0; $i < $colCount; $i++) {
        $xml .= '<w:gridCol w:w="' . ($useTemplateStyle ? 1250 : 0) . '"/>';
    }
    $xml .= '</w:tblGrid>';

    foreach ($tableData as $rowIndex => $row) {
        $xml .= '<w:tr>';
        foreach ($row as $cellIndex => $cellData) {
            $cellValue = is_array($cellData) ? ($cellData['value'] ?? $cellData) : $cellData;
            $bgColor = is_array($cellData) && isset($cellData['bgColor']) ? $cellData['bgColor'] : ($rowIndex == 0 && $useTemplateStyle ? '0070C0' : null);
            $fontColor = is_array($cellData) && isset($cellData['fontColor']) ? $cellData['fontColor'] : ($rowIndex == 0 && $useTemplateStyle ? 'FFFFFF' : ($rowIndex > 0 && $cellIndex == $colCount - 1 && $useTemplateStyle ? '00B050' : null));

            $xml .= '<w:tc>';
            $xml .= '<w:tcPr><w:tcW w:w="' . ($useTemplateStyle ? 1250 : 0) . '" w:type="dxa"/>';
            if ($bgColor) {
                $xml .= "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"$bgColor\"/>";
            }
            $xml .= '</w:tcPr>';
            $xml .= '<w:p><w:pPr><w:jc w:val="' . ($useTemplateStyle ? 'center' : 'left') . '"/></w:pPr><w:r>';
            if ($fontColor) {
                $xml .= "<w:rPr><w:sz w:val=\"" . ($rowIndex == 0 && $useTemplateStyle ? 24 : 22) . "\"/><w:color w:val=\"$fontColor\"/></w:rPr>";
            } elseif ($useTemplateStyle) {
                $xml .= '<w:rPr><w:sz w:val="' . ($rowIndex == 0 ? 24 : 22) . '"/></w:rPr>';
            }
            $xml .= '<w:t>' . htmlspecialchars((string)$cellValue, ENT_QUOTES, 'UTF-8') . '</w:t>';
            $xml .= '</w:r></w:p>';
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
    'companyName' => 'Jersey Engineering Solutions Pvt Ltd.',
    'CaseNo' => 'NADAL\Jersey\001\2023',
    'fullName' => 'Karthikeyan A',
    'reportStatus' => 'Clear',
    'dob' => '1990-01-01',
    'fatherName' => 'Arumugam',
    'requestDate' => '2025-08-20',
    'caseId' => 'CID-98765',
    'reportDate' => '2025-08-21',
    'arg_ta(i)' => [
        [
            ['value' => '1', 'bgColor' => 'FFFF00', 'fontColor' => '000000'],
            ['value' => 'Address Verification', 'bgColor' => 'FFFF00', 'fontColor' => '000000'],
            ['value' => 'Verified', 'bgColor' => '00FF00', 'fontColor' => '000000'],
            ['value' => 'Green', 'bgColor' => '00FF00', 'fontColor' => '000000']
        ],
        [
            ['value' => '2', 'bgColor' => 'FFFFFF', 'fontColor' => 'FF0000'],
            ['value' => 'Educational Verification', 'bgColor' => 'FFFFFF', 'fontColor' => 'FF0000'],
            ['value' => 'Verified', 'bgColor' => '00FF00', 'fontColor' => '000000'],
            ['value' => 'Green', 'bgColor' => '00FF00', 'fontColor' => '000000']
        ]
    ],
    'arg_t(j)' => [
        [
            ['value' => 'fh A\43 Selvamaruthur Thisayanvelai Tirunelveli Tamil Nadu India', 'bgColor' => 'FFFF00', 'fontColor' => '000000'],
            ['value' => 'UTA', 'bgColor' => '00FF00', 'fontColor' => '000000']
        ],[
            ['value' => 'fh A\43 Selvamaruthur Thisayanvelai Tirunelveli Tamil Nadu India', 'bgColor' => 'FFFF00', 'fontColor' => '000000'],
            ['value' => 'UTA', 'bgColor' => '00FF00', 'fontColor' => '000000']
        ],[
            ['value' => 'fh A\43 Selvamaruthur Thisayanvelai Tirunelveli Tamil Nadu India', 'bgColor' => 'FFFF00', 'fontColor' => '000000'],
            ['value' => 'UTA', 'bgColor' => '00FF00', 'fontColor' => '000000']
        ]
    ],
    'addressAsAadhar' => '112/1,Arijan Colony, Eraiyadikal, Chengalakurichi,
    Dohnavur, Tirunelveli,
    Tamil Nadu - 627 612, India.',
    'modeOfVerify' => 'Physically visited no one is available.',
    'addressAsPresent' => '100, G3,Arputham Complex, 
    Near Vivekanda School, Eruvadi,
    Vallioor road, Anna Nagar, Vadakku Vallioor.',
    'resName' => 'anithaa',
    'relSubject' => 'Wife',
    'resNumber' => '9876543210',
    'resResidence' => 'Rented',
    'resStay'=>'02 Years',
    'mode' => 'Physical',
    'remark' => 'Verified',
    'add_img(i)'=>['1' => [
        'image' => 'C:\xampp\htdocs\test\pdfgenback\openxml\image1.jpg',
        'width' => 300,
        'height' => 300
    ],
    '2' => [
        'image' => 'C:\xampp\htdocs\test\pdfgenback\openxml\image2.jpg',
        'width' => 600,
        'height' => 400
    ],],
    'canDegeree' => 'Electrical and Electronics Engineering ',
    'collegeLoc' => '#6, Tiruchendur Road, Vallioor, 
                    Tirunelveli - 627 117.',
    'yearOfpassed' => '2010',
    'collegeName' => 'Vallioor Engineering College',
    'universityname' => 'Anna University',
    'addComm' => 'NADAL has checked the Fake University List and the University Grant Commission website. The University is NOT listed as fake or unaccredited.',
    'verifierName' => 'R. Kalai Selvi',
    'verifierDesignation' => 'HOD',
    'deptName' => 'Department of Electrical and Electronics Engineering',
    'edu_img(i)'=>['1' => [
        'image' => 'C:\xampp\htdocs\test\pdfgenback\openxml\image1.jpg',
        'width' => 300,
        'height' => 300
    ]],
    'criDurationOfRecordsCheck' => '5 Years',
    'criRemark' => 'No records found on the judiciary sites of Sessions Court, High Court, Magistrate Court and Civil Court for the address provided.',
    'result' => 'All Verified 100% Clear',
    'addDOV' => '2025-08-22',
    'eduDOV' => '2006-05-15',
    'eduTOV' => '12.19 PM',
    'criDOV' => '2025-08-22',
    'criSD' => '2025-08-22',
    'criED' => '2025-08-22',
    'successStatus' => '✓',
    'failStatus' => '✖',
];
$template = 'input.docx';
$output = 'output.docx';

// Check if template file exists
if (!file_exists($template)) {
    die("❌ Template file not found: $template\n");
}

// Check if template is readable
if (!is_readable($template)) {
    die("❌ Template file is not readable: $template\n");
}

// Check if output directory is writable
$outputDir = dirname($output);
if (!is_writable($outputDir)) {
    die("❌ Output directory is not writable: $outputDir\n");
}

echo "📄 Template found: $template\n";
echo "🔄 Processing...\n";

replaceDocxPlaceholdersSmart($template, $output, $values);

echo "✅ Created $output\n";
?>
