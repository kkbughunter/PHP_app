<?php


/**
 * Replace %*varName*% placeholders in a DOCX, even when broken across runs,
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

    // Process the main document.xml plus any headers/footers if present.
    $parts = ['word/document.xml'];
    // Add headers/footers if they exist
    for ($i = 1; $i <= 8; $i++) {
        foreach (["word/header$i.xml", "word/footer$i.xml"] as $p) {
            if ($zip->locateName($p) !== false) $parts[] = $p;
        }
    }

    foreach ($parts as $partPath) {
        $index = $zip->locateName($partPath);
        if ($index === false) continue;

        $xmlData = $zip->getFromIndex($index);
        $newXml  = replaceInWordXml($xmlData, $replacements);

        // Overwrite the entry
        $zip->deleteName($partPath);
        $zip->addFromString($partPath, $newXml);
    }

    $zip->close();
}


/**
 * Core logic: walk each paragraph, scan its w:t nodes in order,
 * find %*...*% (across nodes), and replace with provided value.
 */
function replaceInWordXml(string $xmlData, array $replacements): string
{
    $doc = new DOMDocument();
    // Preserve whitespace as Word cares about it in places; do not reformat.
    $doc->preserveWhiteSpace = true;
    $doc->formatOutput = false;
    $doc->loadXML($xmlData);

    $xp = new DOMXPath($doc);
    $xp->registerNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');

    // Process each paragraph independently; placeholders won't cross paragraphs.
    $paragraphs = $xp->query('//w:p');
    foreach ($paragraphs as $p) {
        // Snapshot the <w:t> nodes into a PHP array to avoid live NodeList issues.
        $tNodesList = $xp->query('.//w:r/w:t', $p);
        $tNodes = [];
        foreach ($tNodesList as $tn) { $tNodes[] = $tn; }

        $i = 0;
        while ($i < count($tNodes)) {
            /** @var DOMElement $startNode */
            $startNode = $tNodes[$i];
            $startText = $startNode->nodeValue;
            $posStart  = strpos($startText, '%*');

            if ($posStart === false) {
                $i++;
                continue;
            }

            // Found a start marker in this node.
            $startIndex  = $i;
            $startOffset = $posStart;

            // Gather text until end marker *% is found (could be same node or later).
            $varText = '';
            $endFound = false;

            // First, remainder of start node after '%*'
            $afterStart = substr($startText, $startOffset + 2);

            // Check if end also exists in this same node
            $posEndSame = strpos($afterStart, '*%');
            if ($posEndSame !== false) {
                // Same-node placeholder
                $varRaw    = substr($afterStart, 0, $posEndSame);
                $trailing  = substr($afterStart, $posEndSame + 2);
                $leading   = substr($startText, 0, $startOffset);

                $varName = preg_replace('/\s+/', '', $varRaw);
                $replacement = array_key_exists($varName, $replacements) ? (string)$replacements[$varName] : null;

                if ($replacement !== null) {
                    $startNode->nodeValue = $leading . $replacement . $trailing;
                } else {
                    // No replacement provided: remove markers but keep inner text as-is
                    $startNode->nodeValue = $leading . $varRaw . $trailing;
                }

                // Continue scanning the same node from after what we just wrote.
                // Move $i forward to re-scan from this node in case there are more placeholders later in the node.
                continue;
            }

            // Multi-node placeholder
            $varText .= $afterStart;
            $endIndex = $startIndex;
            $endOffset = null;

            for ($j = $startIndex + 1; $j < count($tNodes); $j++) {
                $endCandidate = $tNodes[$j]->nodeValue;
                $posEnd = strpos($endCandidate, '*%');
                if ($posEnd === false) {
                    $varText .= $endCandidate;
                    continue;
                }
                // Found end
                $varText   .= substr($endCandidate, 0, $posEnd);
                $endIndex   = $j;
                $endOffset  = $posEnd;
                $endFound   = true;
                break;
            }

            if (!$endFound) {
                // Unbalanced start marker: nothing to do; move on.
                $i++;
                continue;
            }

            // Compute leading/trailing text for the boundary nodes
            $leadingStart = substr($startText, 0, $startOffset);
            $endNodeText  = $tNodes[$endIndex]->nodeValue;
            $trailingEnd  = substr($endNodeText, $endOffset + 2); // after '*%'

            $varName = preg_replace('/\s+/', '', $varText);
            $replacement = array_key_exists($varName, $replacements) ? (string)$replacements[$varName] : null;

            if ($replacement !== null) {
                // Put leading + replacement into the start node
                $tNodes[$startIndex]->nodeValue = $leadingStart . $replacement;
            } else {
                // No replacement: strip markers but keep original inner text
                $tNodes[$startIndex]->nodeValue = $leadingStart . $varText;
            }

            // Clear all middle nodes completely (they held parts of the placeholder)
            for ($k = $startIndex + 1; $k < $endIndex; $k++) {
                $tNodes[$k]->nodeValue = '';
            }

            // Put only trailing text into the end node
            $tNodes[$endIndex]->nodeValue = $trailingEnd;

            // Resume scanning **after** the end node
            $i = $endIndex;
        }
    }

    // Return the updated XML
    return $doc->saveXML();
}

// ------------------------------ Example usage ------------------------------
// IMPORTANT: Your template must contain placeholders like %*employerName*% (no spaces).
$template = __DIR__ . '/example.docx';   // <-- change to your template path
$output   = __DIR__ . '/output.docx';

$values = [
    'variable' => 'Hello!....',
    'employerName' => 'Google Inc.',
    'CaseID'       => 'BGV123456',
    'cName'        => 'John Doe',
    'reportStatus' => 'Clear',
    'dob'          => '1990-01-01',
    'reuestDate'   => '2025-08-20', // matches your template’s typo
    'cID'          => 'CID-98765',
    'reportDate'   => '2025-08-21',
];

replaceDocxPlaceholdersSmart($template, $output, $values);

echo "✅ Created $output\n";