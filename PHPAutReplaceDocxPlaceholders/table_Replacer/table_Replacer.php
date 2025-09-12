<?php

/**
 * Replace %*varName*% placeholders in a DOCX for table row updates, where placeholders like %*tableA1*%, %*tableA2*% etc.
 * indicate a template row. Replaces the first data row into the template row (including applying bgColor/fontColor),
 * then appends new rows copying the template structure with data and styling.
 * Supports general prefixes (e.g., 'tableA' maps to 'tableA(i)' or 'tableA(j)' if present in replacements).
 * Writes to a new output file.
 *
 * Requirements: ext-dom, ext-zip
 */
function replaceDocxTableRows(string $inputDocx, string $outputDocx, array $replacements): void
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

    // Process the main document.xml plus any headers/footers
    $parts = ['word/document.xml'];
    for ($i = 1; $i <= 8; $i++) {
        foreach (["word/header$i.xml", "word/footer$i.xml"] as $p) {
            if ($zip->locateName($p) !== false) $parts[] = $p;
        }
    }

    foreach ($parts as $partPath) {
        $index = $zip->locateName($partPath);
        if ($index === false) continue;

        $xmlData = $zip->getFromIndex($index);
        $newXml = replaceTableRowsInWordXml($xmlData, $replacements);

        // Overwrite the XML part
        $zip->deleteName($partPath);
        $zip->addFromString($partPath, $newXml);
    }

    $zip->close();
}

/**
 * Core logic: process tables for row-specific placeholder replacements with dynamic prefix detection.
 */
function replaceTableRowsInWordXml(string $xmlData, array $replacements): string
{
    $doc = new DOMDocument();
    $doc->preserveWhiteSpace = true;
    $doc->formatOutput = false;
    if (!$doc->loadXML($xmlData)) {
        throw new RuntimeException("Failed to load XML data");
    }

    $xp = new DOMXPath($doc);
    $xp->registerNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');

    // Process tables for row-specific placeholder replacements
    $tables = $xp->query('//w:tbl');
    foreach ($tables as $table) {
        processTableRows($table, $replacements, $xp, $doc);
    }

    return $doc->saveXML();
}

/**
 * Process table rows to replace placeholders and add new rows with preserved or user-specified styling.
 * Dynamically detects prefix from placeholders (e.g., 'tableA' from 'tableA1') and matches to keys like 'tableA(i)'.
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
            // Identify the prefix from the first placeholder (e.g., 'tableA' from 'tableA1')
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

            // Find matching key in replacements (e.g., 'tableA(i)' or 'tableB(j)')
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
                    if (isset($placeholders[$cellIndex])) {
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
                                $colIndex = 0; // Special case for %*1*% in tableA
                            } else {
                                continue;
                            }
                            $cellValue = isset($firstRowData[$colIndex]) ? (string)(is_array($firstRowData[$colIndex]) ? ($firstRowData[$colIndex]['value'] ?? $firstRowData[$colIndex]) : $firstRowData[$colIndex]) : '';

                            // Clear existing runs and insert new text
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
                                $newText = $doc->createElementNS($namespace, 'w:t');
                                $newText->nodeValue = htmlspecialchars($cellValue, ENT_QUOTES, 'UTF-8');
                                $newRun->appendChild($newText);
                                $para->appendChild($newRun);
                            }
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

// ------------------------------ Example usage ------------------------------

$result = [
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
    'tableC(i)' => [
        [
            ['value' => '1', 'bgColor' => 'FFFF00', 'fontColor' => '000000'],
            ['value' => 'Verified', 'bgColor' => '00FF00', 'fontColor' => '000000'],
            ['value' => 'Green', 'bgColor' => '00FF00', 'fontColor' => '000000']
        ],
        [
            ['value' => '2', 'bgColor' => 'FFFFFF', 'fontColor' => 'FF0000'],
            ['value' => 'Educational Verification', 'bgColor' => 'FFFFFF', 'fontColor' => 'FF0000'],
            ['value' => 'Green', 'bgColor' => '00FF00', 'fontColor' => '000000']
        ]
    ],
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
echo "🔄 Processing tables...\n";

replaceDocxTableRows($template, $output, $result);

echo "✅ Created $output\n";
?>


    