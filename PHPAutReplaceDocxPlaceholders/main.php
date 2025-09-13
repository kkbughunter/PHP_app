
<?php

/**
 * Replace %*varName*% placeholders in a DOCX, including images and tables, even when broken across runs,
 * and handle table row placeholders (e.g., tableA1, tableA2), adding new rows with preserved styling.
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
 * Core logic: first process tables for row replacements, then walk each paragraph for text, image, or table replacements.
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

    // First, process tables for row-specific placeholder replacements
    $tables = $xp->query('//w:tbl');
    foreach ($tables as $table) {
        processTableRows($table, $replacements, $xp, $doc);
    }

    // Then, process paragraphs for text, image, or full table replacements
    $paragraphs = $xp->query('//w:p');
    foreach ($paragraphs as $p) {
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

            $startIndex = $i;
            $startOffset = $posStart;
            $varText = '';
            $endFound = false;
            $endIndex = $startIndex;
            $endOffset = null;

            $afterStart = substr($startText, $startOffset + 2);
            $posEndSame = strpos($afterStart, '*%');
            if ($posEndSame !== false) {
                // Same-node placeholder
                $varRaw = substr($afterStart, 0, $posEndSame);
                $trailing = substr($afterStart, $posEndSame + 2);
                $leading = substr($startText, 0, $startOffset);
                $varName = preg_replace('/\s+/', '', $varRaw);

                if (preg_match('/^(.+)\(i\)$/', $varName, $matches)) {
                    $baseImageName = $matches[1];
                    if (isset($replacements[$varName]) && is_array($replacements[$varName])) {
                        $imageXmlParts = [];
                        foreach ($replacements[$varName] as $imageKey => $imageData) {
                            if (is_array($imageData) && isset($imageData['image'])) {
                                $ext = strtolower(pathinfo($imageData['image'], PATHINFO_EXTENSION));
                                $imageExtensions[$ext] = true;
                                $rId = "rIdImg" . $imageCounter++;
                                if (!isset($imageCounters[$baseImageName])) {
                                    $imageCounters[$baseImageName] = 1;
                                }
                                $imageName = $baseImageName . "-" . $imageCounters[$baseImageName]++ . "." . $ext;
                                $imageRelationships[$rId] = [
                                    'path' => $imageData['image'],
                                    'name' => $imageName
                                ];
                                $relNode = $relsDoc->createElementNS('http://schemas.openxmlformats.org/package/2006/relationships', 'r:Relationship');
                                $relNode->setAttribute('Id', $rId);
                                $relNode->setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
                                $relNode->setAttribute('Target', "media/$imageName");
                                $relsDoc->documentElement->appendChild($relNode);
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
                            if ($p->parentNode) {
                                $p->parentNode->replaceChild($imageFragment, $p);
                            }
                            break;
                        }
                    }
                } elseif (isset($replacements[$varName]) && is_array($replacements[$varName]) && isset($replacements[$varName]['image'])) {
                    $imageData = $replacements[$varName];
                    $ext = strtolower(pathinfo($imageData['image'], PATHINFO_EXTENSION));
                    $imageExtensions[$ext] = true;
                    $rId = "rIdImg" . $imageCounter++;
                    $imageName = "image" . $imageCounter . "." . $ext;
                    $imageRelationships[$rId] = [
                        'path' => $imageData['image'],
                        'name' => $imageName
                    ];
                    $relNode = $relsDoc->createElementNS('http://schemas.openxmlformats.org/package/2006/relationships', 'r:Relationship');
                    $relNode->setAttribute('Id', $rId);
                    $relNode->setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
                    $relNode->setAttribute('Target', "media/$imageName");
                    $relsDoc->documentElement->appendChild($relNode);
                    $imageXml = createImageXml($rId, $imageName, $imageData['width'] ?? 300, $imageData['height'] ?? 300);
                    $imageFragment = $doc->createDocumentFragment();
                    if (!$imageFragment->appendXML($imageXml)) {
                        throw new RuntimeException("Failed to append image XML for $varName");
                    }
                    if ($startRun->parentNode) {
                        $startRun->parentNode->replaceChild($imageFragment, $startRun);
                    }
                } elseif (isset($replacements[$varName]) && is_array($replacements[$varName]) && is_array($replacements[$varName][0])) {
                    $tableData = $replacements[$varName];
                    $tableXml = createTableXml($tableData, true);
                    $tableFragment = $doc->createDocumentFragment();
                    if (!$tableFragment->appendXML($tableXml)) {
                        throw new RuntimeException("Failed to append table XML for $varName");
                    }
                    if ($p->parentNode) {
                        $p->parentNode->replaceChild($tableFragment, $p);
                    }
                    break;
                } else {
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

            if (preg_match('/^(.+)\(i\)$/', $varName, $matches)) {
                $baseImageName = $matches[1];
                if (isset($replacements[$varName]) && is_array($replacements[$varName])) {
                    $imageXmlParts = [];
                    foreach ($replacements[$varName] as $imageKey => $imageData) {
                        if (is_array($imageData) && isset($imageData['image'])) {
                            $ext = strtolower(pathinfo($imageData['image'], PATHINFO_EXTENSION));
                            $imageExtensions[$ext] = true;
                            $rId = "rIdImg" . $imageCounter++;
                            if (!isset($imageCounters[$baseImageName])) {
                                $imageCounters[$baseImageName] = 1;
                            }
                            $imageName = $baseImageName . "-" . $imageCounters[$baseImageName]++ . "." . $ext;
                            $imageRelationships[$rId] = [
                                'path' => $imageData['image'],
                                'name' => $imageName
                            ];
                            $relNode = $relsDoc->createElementNS('http://schemas.openxmlformats.org/package/2006/relationships', 'r:Relationship');
                            $relNode->setAttribute('Id', $rId);
                            $relNode->setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
                            $relNode->setAttribute('Target', "media/$imageName");
                            $relsDoc->documentElement->appendChild($relNode);
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
                        if ($p->parentNode) {
                            $p->parentNode->replaceChild($imageFragment, $p);
                        }
                        break;
                    }
                }
            } elseif (isset($replacements[$varName]) && is_array($replacements[$varName]) && isset($replacements[$varName]['image'])) {
                $imageData = $replacements[$varName];
                $ext = strtolower(pathinfo($imageData['image'], PATHINFO_EXTENSION));
                $imageExtensions[$ext] = true;
                $rId = "rIdImg" . $imageCounter++;
                $imageName = "image" . $imageCounter . "." . $ext;
                $imageRelationships[$rId] = [
                    'path' => $imageData['image'],
                    'name' => $imageName
                ];
                $relNode = $relsDoc->createElementNS('http://schemas.openxmlformats.org/package/2006/relationships', 'r:Relationship');
                $relNode->setAttribute('Id', $rId);
                $relNode->setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
                $relNode->setAttribute('Target', "media/$imageName");
                $relsDoc->documentElement->appendChild($relNode);
                $imageXml = createImageXml($rId, $imageName, $imageData['width'] ?? 300, $imageData['height'] ?? 300);
                $imageFragment = $doc->createDocumentFragment();
                if (!$imageFragment->appendXML($imageXml)) {
                    throw new RuntimeException("Failed to append image XML for $varName");
                }
                if ($startRun->parentNode) {
                    $startRun->parentNode->replaceChild($imageFragment, $startRun);
                }
                for ($k = $startIndex + 1; $k < $endIndex; $k++) {
                    if ($runs[$k]['run']->parentNode) {
                        $runs[$k]['run']->parentNode->removeChild($runs[$k]['run']);
                    }
                }
                if ($endTextNode) {
                    $endTextNode->nodeValue = $trailingEnd;
                }
            } elseif (isset($replacements[$varName]) && is_array($replacements[$varName]) && is_array($replacements[$varName][0])) {
                $tableData = $replacements[$varName];
                $tableXml = createTableXml($tableData, true);
                $tableFragment = $doc->createDocumentFragment();
                if (!$tableFragment->appendXML($tableXml)) {
                    throw new RuntimeException("Failed to append table XML for $varName");
                }
                if ($p->parentNode) {
                    $p->parentNode->replaceChild($tableFragment, $p);
                }
                break;
            } else {
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
 * Process table rows to replace placeholders and add new rows with preserved or user-specified styling.
 * Dynamically detects prefix from placeholders (e.g., 'tableC' from 'tableC1') and matches to keys like 'tableC(k)'.
 */
function processTableRows(DOMElement $table, array $replacements, DOMXPath $xp, DOMDocument $doc): void
{
    $rows = $xp->query('w:tr', $table);
    $rowArray = [];
    foreach ($rows as $row) {
        $rowArray[] = $row;
    }

    $namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

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
            $run = $para ? $xp->query('w:r', $para)->item(0) : null;
            $rPr = $run ? $xp->query('w:rPr', $run)->item(0) : null;
            $runProperties[$cellIndex] = $rPr ? $doc->saveXML($rPr) : '';
        }

        // Check each cell for placeholders by concatenating all text in the paragraph
        foreach ($cells as $cellIndex => $cell) {
            $para = $xp->query('w:p', $cell)->item(0);
            if (!$para) continue;
            $textNodes = $xp->query('.//w:t', $para);
            $cellText = '';
            foreach ($textNodes as $tNode) {
                $cellText .= $tNode->nodeValue;
            }
            preg_match_all('/%\*[^*]+\*%/', $cellText, $matches);
            if (!empty($matches[0])) {
                $hasPlaceholders = true;
                $placeholders[$cellIndex] = $matches[0];
            }
        }

        if ($hasPlaceholders) {
            // Identify the prefix from the first placeholder (e.g., 'tableC' from 'tableC1')
            $firstPlaceholder = reset($placeholders)[0] ?? '';
            $firstVarName = '';
            if (preg_match('/%\*([^*]+)\*%/', $firstPlaceholder, $m)) {
                $firstVarName = trim($m[1]);
            }
            $prefix = '';
            if (preg_match('/^(.+?)(\d+)$/', $firstVarName, $prefixMatches)) {
                $prefix = $prefixMatches[1];
            } elseif (preg_match('/^\d+$/', $firstVarName)) {
                $prefix = 'tableA'; // Handle special case for %*1*% in tableA
            } else {
                continue; // Not a valid numbered prefix
            }

            // Find matching key in replacements (e.g., 'tableC(k)' or 'tableD(l)')
            $possibleKeys = [];
            foreach (array_keys($replacements) as $key) {
                if (preg_match('/^' . preg_quote($prefix, '/') . '\(\w+\)$/', $key)) {
                    $possibleKeys[] = $key;
                }
            }
            if (count($possibleKeys) !== 1) {
                continue; // No unique match
            }
            $baseKey = $possibleKeys[0];

            if (isset($replacements[$baseKey]) && is_array($replacements[$baseKey]) && !empty($replacements[$baseKey]) && is_array($replacements[$baseKey][0])) {
                $tableData = $replacements[$baseKey];
                $firstRowData = array_shift($tableData);

                // Replace placeholders in the current row with the first row of data
                foreach ($cells as $cellIndex => $cell) {
                    if (isset($placeholders[$cellIndex]) && !empty($placeholders[$cellIndex])) {
                        $para = $xp->query('w:p', $cell)->item(0);
                        if ($para) {
                            $runs = $xp->query('w:r', $para);
                            foreach ($runs as $run) {
                                $run->parentNode->removeChild($run);
                            }
                            $newRun = $doc->createElementNS($namespace, 'w:r');
                            if ($runProperties[$cellIndex]) {
                                $rPrDoc = new DOMDocument();
                                $rPrDoc->loadXML('<root xmlns:w="' . $namespace . '">' . $runProperties[$cellIndex] . '</root>');
                                $rPrNode = $rPrDoc->documentElement->firstChild;
                                $importedRPr = $doc->importNode($rPrNode, true);
                                $newRun->appendChild($importedRPr);
                            }
                            $cellValue = '';
                            // Handle multiple placeholders per cell if needed, but take the last matching one for simplicity
                            foreach ($placeholders[$cellIndex] as $placeholder) {
                                $colVarName = '';
                                if (preg_match('/%\*([^*]+)\*%/', $placeholder, $m)) {
                                    $colVarName = trim($m[1]);
                                }
                                $colIndex = -1;
                                if (preg_match('/^' . preg_quote($prefix, '/') . '(\d+)$/', $colVarName, $colMatches)) {
                                    $colNum = (int)$colMatches[1];
                                    $colIndex = $colNum - 1;
                                } elseif ($colVarName === '1' && $prefix === 'tableA') {
                                    $colIndex = 0; // Special case for %*1*%
                                } else {
                                    continue;
                                }
                                if ($colIndex >= 0) {
                                    $cellValue = isset($firstRowData[$colIndex]) ? (string)(is_array($firstRowData[$colIndex]) ? ($firstRowData[$colIndex]['value'] ?? $firstRowData[$colIndex]) : $firstRowData[$colIndex]) : '';
                                }
                            }
                            $newText = $doc->createElementNS($namespace, 'w:t');
                            $newText->nodeValue = htmlspecialchars($cellValue, ENT_QUOTES, 'UTF-8');
                            $newRun->appendChild($newText);
                            $para->appendChild($newRun);
                        }
                    }
                }

                // Apply styling (bgColor, fontColor) to the current (first) row
                foreach ($cells as $cellIndex => $cell) {
                    if (isset($firstRowData[$cellIndex]) && is_array($firstRowData[$cellIndex])) {
                        $cellData = $firstRowData[$cellIndex];
                        $bgColor = $cellData['bgColor'] ?? null;
                        $fontColor = $cellData['fontColor'] ?? null;

                        // Modify tcPr for bgColor
                        $tcPr = $xp->query('w:tcPr', $cell)->item(0);
                        if ($tcPr && $bgColor) {
                            $shd = $xp->query('w:shd', $tcPr)->item(0);
                            if ($shd) {
                                $shd->setAttribute('w:fill', $bgColor);
                            } else {
                                $newShd = $doc->createElementNS($namespace, 'w:shd');
                                $newShd->setAttribute('w:val', 'clear');
                                $newShd->setAttribute('w:color', 'auto');
                                $newShd->setAttribute('w:fill', $bgColor);
                                $tcPr->appendChild($newShd);
                            }
                        }

                        // Modify rPr for fontColor
                        $para = $xp->query('w:p', $cell)->item(0);
                        if ($para && $fontColor) {
                            $run = $xp->query('w:r', $para)->item(0);
                            if ($run) {
                                $rPr = $xp->query('w:rPr', $run)->item(0);
                                if ($rPr) {
                                    $color = $xp->query('w:color', $rPr)->item(0);
                                    if ($color) {
                                        $color->setAttribute('w:val', $fontColor);
                                    } else {
                                        $newColor = $doc->createElementNS($namespace, 'w:color');
                                        $newColor->setAttribute('w:val', $fontColor);
                                        $rPr->appendChild($newColor);
                                    }
                                } else {
                                    $newRPr = $doc->createElementNS($namespace, 'w:rPr');
                                    $newColor = $doc->createElementNS($namespace, 'w:color');
                                    $newColor->setAttribute('w:val', $fontColor);
                                    $newRPr->appendChild($newColor);
                                    $run->insertBefore($newRPr, $run->firstChild);
                                }
                            }
                        }
                    }
                }

                // Add new rows with preserved styling from the template row or user-specified colors
                foreach ($tableData as $rowData) {
                    $newRowXml = '<w:tr xmlns:w="' . $namespace . '">';
                    foreach ($rowData as $cellIndex => $cellData) {
                        $cellValue = is_array($cellData) ? ($cellData['value'] ?? $cellData) : $cellData;
                        $bgColor = is_array($cellData) && isset($cellData['bgColor']) ? $cellData['bgColor'] : null;
                        $fontColor = is_array($cellData) && isset($cellData['fontColor']) ? $cellData['fontColor'] : null;

                        $newRowXml .= '<w:tc>';

                        // Create/modify tcPr XML for bgColor
                        $tcPrXml = $cellProperties[$cellIndex] ?? '<w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>';
                        if ($bgColor) {
                            $tcPrDoc = new DOMDocument();
                            $tcPrDoc->loadXML('<root xmlns:w="' . $namespace . '">' . $tcPrXml . '</root>');
                            $tcPrXp = new DOMXPath($tcPrDoc);
                            $tcPrXp->registerNamespace('w', $namespace);
                            $shdNode = $tcPrXp->query('//w:shd')->item(0);
                            if ($shdNode) {
                                $shdNode->setAttribute('w:fill', $bgColor);
                            } else {
                                $tcPrNode = $tcPrXp->query('//w:tcPr')->item(0);
                                if ($tcPrNode) {
                                    $newShdNode = $tcPrDoc->createElementNS($namespace, 'w:shd');
                                    $newShdNode->setAttribute('w:val', 'clear');
                                    $newShdNode->setAttribute('w:color', 'auto');
                                    $newShdNode->setAttribute('w:fill', $bgColor);
                                    $tcPrNode->appendChild($newShdNode);
                                }
                            }
                            $tcPrXml = $tcPrDoc->saveXML($tcPrDoc->documentElement->firstChild);
                        }
                        $newRowXml .= $tcPrXml;

                        $newRowXml .= '<w:p>';
                        if (!empty($paraProperties[$cellIndex])) {
                            $newRowXml .= $paraProperties[$cellIndex];
                        }
                        $newRowXml .= '<w:r>';
                        // Create/modify run XML for fontColor
                        $runXml = $runProperties[$cellIndex] ?? '';
                        if ($fontColor && $runXml) {
                            $runDoc = new DOMDocument();
                            $runDoc->loadXML('<root xmlns:w="' . $namespace . '">' . $runXml . '</root>');
                            $runXp = new DOMXPath($runDoc);
                            $runXp->registerNamespace('w', $namespace);
                            $colorNode = $runXp->query('//w:color')->item(0);
                            if ($colorNode) {
                                $colorNode->setAttribute('w:val', $fontColor);
                            } else {
                                $rPrNode = $runXp->query('//w:rPr')->item(0);
                                if ($rPrNode) {
                                    $newColorNode = $runDoc->createElementNS($namespace, 'w:color');
                                    $newColorNode->setAttribute('w:val', $fontColor);
                                    $rPrNode->appendChild($newColorNode);
                                }
                            }
                            $runXml = $runDoc->saveXML($runDoc->documentElement->firstChild);
                        } elseif ($fontColor) {
                            $runXml = '<w:rPr xmlns:w="' . $namespace . '"><w:color w:val="' . $fontColor . '"/><w:sz w:val="28"/></w:rPr>';
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
    'tableA(i)' => [
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
    'tableB(j)' => [
        [
            ['value' => 'fh A\43 Selvamaruthur Thisayanvelai Tirunelveli Tamil Nadu India', 'bgColor' => 'FFFF00', 'fontColor' => '000000'],
            ['value' => 'UTA', 'bgColor' => '00FF00', 'fontColor' => '000000']
        ],
        [
            ['value' => 'fh A\43 Selvamaruthur Thisayanvelai Tirunelveli Tamil Nadu India', 'bgColor' => 'FFFF00', 'fontColor' => '000000'],
            ['value' => 'UTA', 'bgColor' => '00FF00', 'fontColor' => '000000']
        ],
        [
            ['value' => 'fh A\43 Selvamaruthur Thisayanvelai Tirunelveli Tamil Nadu India', 'bgColor' => 'FFFF00', 'fontColor' => '000000'],
            ['value' => 'UTA', 'bgColor' => '00FF00', 'fontColor' => '000000']
        ]
    ],
    'tableC(k)' => [
        [
            ['value' => 'table C content A1', 'bgColor' => 'FFFF00', 'fontColor' => '000000'],
            ['value' => 'table C content B1', 'bgColor' => '00FF00', 'fontColor' => '000000']
        ],
        [
            ['value' => 'table C content A2', 'bgColor' => 'FFFF00', 'fontColor' => '000000'],
            ['value' => 'table C content B2', 'bgColor' => '00FF00', 'fontColor' => '000000']
        ],
        [
            ['value' => 'table C content A3', 'bgColor' => 'FFFF00', 'fontColor' => '000000'],
            ['value' => 'table C content B3', 'bgColor' => '00FF00', 'fontColor' => '000000']
        ]
    ],
    // Example for tableD - add this to $values to support it
    // 'tableD(l)' => [
    //     [
    //         ['value' => 'table D content 1', 'bgColor' => 'FF0000', 'fontColor' => 'FFFFFF'],
    //         ['value' => 'table D content 2', 'bgColor' => '0000FF', 'fontColor' => 'FFFFFF'],
    //         // Add more columns as needed
    //     ],
    //     // Add more rows as needed
    // ],
    'addressAsAadhar' => '112/1,Arijan Colony, Eraiyadikal, Chengalakurichi, Dohnavur, Tirunelveli, Tamil Nadu - 627 612, India.',
    'modeOfVerify' => 'Physically visited no one is available.',
    'addressAsPresent' => '100, G3,Arputham Complex, Near Vivekanda School, Eruvadi, Vallioor road, Anna Nagar, Vadakku Vallioor.',
    'resName' => 'anithaa',
    'relSubject' => 'Wife',
    'resNumber' => '9876543210',
    'resResidence' => 'Rented',
    'resStay' => '02 Years',
    'mode' => 'Physical',
    'remark' => 'Verified',
    'add_img(i)' => [
        '1' => [
            'image' => 'C:\xampp\htdocs\test\pdfgenback\openxml\image1.jpg',
            'width' => 300,
            'height' => 300
        ],
        '2' => [
            'image' => 'C:\xampp\htdocs\test\pdfgenback\openxml\image2.jpg',
            'width' => 600,
            'height' => 400
        ],
    ],
    'canDegeree' => 'Electrical and Electronics Engineering ',
    'collegeLoc' => '#6, Tiruchendur Road, Vallioor, Tirunelveli - 627 117.',
    'yearOfpassed' => '2010',
    'collegeName' => 'Vallioor Engineering College',
    'universityname' => 'Anna University',
    'addComm' => 'NADAL has checked the Fake University List and the University Grant Commission website. The University is NOT listed as fake or unaccredited.',
    'verifierName' => 'R. Kalai Selvi',
    'verifierDesignation' => 'HOD',
    'deptName' => 'Department of Electrical and Electronics Engineering',
    'edu_img(i)' => [
        '1' => [
            'image' => 'C:\xampp\htdocs\test\pdfgenback\openxml\image1.jpg',
            'width' => 300,
            'height' => 300
        ]
    ],
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
