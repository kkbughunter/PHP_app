<?php
/**
 * Replace placeholders in a DOCX template with text and images
 */

function replaceDocxPlaceholders($templatePath, $outputPath, $replacements) {
    $zip = new ZipArchive;
    if ($zip->open($templatePath) !== TRUE) {
        throw new Exception("Could not open DOCX template: $templatePath");
    }

    // Process each Word part (main doc, headers, footers)
    $parts = ['word/document.xml'];
    for ($i = 1; $i < 10; $i++) {
        if ($zip->locateName("word/header{$i}.xml") !== false) $parts[] = "word/header{$i}.xml";
        if ($zip->locateName("word/footer{$i}.xml") !== false) $parts[] = "word/footer{$i}.xml";
    }

    foreach ($parts as $partPath) {
        $xml = $zip->getFromName($partPath);
        if ($xml === false) continue;

        $doc = new DOMDocument();
        $doc->preserveWhiteSpace = false;
        $doc->formatOutput = true;
        $doc->loadXML($xml);

        // Find all <w:t> nodes
        $xpath = new DOMXPath($doc);
        $xpath->registerNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        $tNodes = $xpath->query("//w:t");

        foreach ($replacements as $placeholder => $replacement) {
            foreach ($tNodes as $node) {
                if (strpos($node->nodeValue, $placeholder) !== false) {
                    if (is_array($replacement) && isset($replacement['image'])) {
                        // --- Handle image replacement ---
                        $imagePath = $replacement['image'];
                        if (!file_exists($imagePath)) {
                            throw new Exception("Image file not found: $imagePath");
                        }

                        $imageName = basename($imagePath);
                        $mediaPath = "word/media/$imageName";

                        // Add image binary to media folder
                        $zip->addFile($imagePath, $mediaPath);

                        // Update relationships
                        $relsPath = dirname($partPath) . '/_rels/' . basename($partPath) . '.rels';
                        $relsXml = $zip->getFromName($relsPath);
                        if ($relsXml === false) {
                            // Create new relationships file if it doesn't exist
                            $relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
                        }

                        $relsDoc = new DOMDocument();
                        $relsDoc->loadXML($relsXml);
                        $relsXPath = new DOMXPath($relsDoc);
                        $relsXPath->registerNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");
                        $existingIds = $relsXPath->query("//r:Relationship/@Id");
                        $usedIds = [];
                        foreach ($existingIds as $id) {
                            $usedIds[] = $id->nodeValue;
                        }

                        // Generate unique rId
                        $rId = 'rId' . rand(1000, 9999);
                        while (in_array($rId, $usedIds)) {
                            $rId = 'rId' . rand(1000, 9999);
                        }

                        $rel = $relsDoc->createElementNS(
                            "http://schemas.openxmlformats.org/package/2006/relationships",
                            'Relationship'
                        );
                        $rel->setAttribute("Id", $rId);
                        $rel->setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                        $rel->setAttribute("Target", "media/$imageName");
                        $relsDoc->documentElement->appendChild($rel);

                        $zip->deleteName($relsPath);
                        $zip->addFromString($relsPath, $relsDoc->saveXML());

                        // Replace placeholder text node with drawing
                        insertImageNode(
                            $doc,
                            $node,
                            $rId,
                            $replacement['width'] ?? 200,
                            $replacement['height'] ?? 200
                        );
                    } else {
                        // --- Handle text replacement ---
                        $node->nodeValue = str_replace($placeholder, $replacement, $node->nodeValue);
                    }
                }
            }
        }

        $zip->deleteName($partPath);
        $zip->addFromString($partPath, $doc->saveXML());
    }

    // Update [Content_Types].xml to include image type
    $contentTypesPath = '[Content_Types].xml';
    $contentTypesXml = $zip->getFromName($contentTypesPath);
    $contentTypesDoc = new DOMDocument();
    $contentTypesDoc->loadXML($contentTypesXml);
    $contentTypesXPath = new DOMXPath($contentTypesDoc);
    $contentTypesXPath->registerNamespace("ct", "http://schemas.openxmlformats.org/package/2006/content-types");
    $imageExtension = pathinfo($replacements['%*img1*%']['image'], PATHINFO_EXTENSION);
    $contentType = 'image/jpeg';
    if ($imageExtension === 'png') {
        $contentType = 'image/png';
    }

    $existingTypes = $contentTypesXPath->query("//ct:Override[@ContentType='$contentType']");
    if ($existingTypes->length === 0) {
        $override = $contentTypesDoc->createElementNS(
            "http://schemas.openxmlformats.org/package/2006/content-types",
            'Override'
        );
        $override->setAttribute("PartName", "/word/media/" . basename($replacements['%*img1*%']['image']));
        $override->setAttribute("ContentType", $contentType);
        $contentTypesDoc->documentElement->appendChild($override);
        $zip->deleteName($contentTypesPath);
        $zip->addFromString($contentTypesPath, $contentTypesDoc->saveXML());
    }

    $zip->close();

    // Copy template to output and reopen to ensure changes are applied
    copy($templatePath, $outputPath);
}

/**
 * Insert an image XML fragment in place of a placeholder node
 */
function insertImageNode(DOMDocument $doc, DOMElement $refNode, string $imageId, int $width, int $height) {
    $ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    $ns_wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    $ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main";
    $ns_pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    $ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    $r = $doc->createElementNS($ns_w, 'w:r');
    $drawing = $doc->createElementNS($ns_w, 'w:drawing');
    $r->appendChild($drawing);

    $inline = $doc->createElementNS($ns_wp, 'wp:inline');
    $inline->setAttribute("distT", "0");
    $inline->setAttribute("distB", "0");
    $inline->setAttribute("distL", "0");
    $inline->setAttribute("distR", "0");
    $drawing->appendChild($inline);

    // Extent (image size in EMUs: 1 pixel = 9525 EMUs)
    $extent = $doc->createElementNS($ns_wp, 'wp:extent');
    $extent->setAttribute("cx", $width * 9525);
    $extent->setAttribute("cy", $height * 9525);
    $inline->appendChild($extent);

    // Document properties
    $docPr = $doc->createElementNS($ns_wp, 'wp:docPr');
    $docPr->setAttribute("id", "1");
    $docPr->setAttribute("name", "Picture 1");
    $inline->appendChild($docPr);

    // Graphic element
    $graphic = $doc->createElementNS($ns_a, 'a:graphic');
    $graphicData = $doc->createElementNS($ns_a, 'a:graphicData');
    $graphicData->setAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture");
    $graphic->appendChild($graphicData);
    $inline->appendChild($graphic);

    // Picture element
    $pic = $doc->createElementNS($ns_pic, 'pic:pic');
    $graphicData->appendChild($pic);

    // Non-visual picture properties
    $nvPicPr = $doc->createElementNS($ns_pic, 'pic:nvPicPr');
    $cNvPr = $doc->createElementNS($ns_pic, 'pic:cNvPr');
    $cNvPr->setAttribute("id", "0");
    $cNvPr->setAttribute("name", "Picture");
    $nvPicPr->appendChild($cNvPr);
    $cNvPicPr = $doc->createElementNS($ns_pic, 'pic:cNvPicPr');
    $nvPicPr->appendChild($cNvPicPr);
    $pic->appendChild($nvPicPr);

    // Blip fill
    $blipFill = $doc->createElementNS($ns_pic, 'pic:blipFill');
    $blip = $doc->createElementNS($ns_a, 'a:blip');
    $blip->setAttributeNS($ns_r, 'r:embed', $imageId);
    $blipFill->appendChild($blip);
    $stretch = $doc->createElementNS($ns_a, 'a:stretch');
    $fillRect = $doc->createElementNS($ns_a, 'a:fillRect');
    $stretch->appendChild($fillRect);
    $blipFill->appendChild($stretch);
    $pic->appendChild($blipFill);

    // Shape properties
    $spPr = $doc->createElementNS($ns_pic, 'pic:spPr');
    $xfrm = $doc->createElementNS($ns_a, 'a:xfrm');
    $off = $doc->createElementNS($ns_a, 'a:off');
    $off->setAttribute("x", "0");
    $off->setAttribute("y", "0");
    $xfrm->appendChild($off);
    $ext = $doc->createElementNS($ns_a, 'a:ext');
    $ext->setAttribute("cx", $width * 9525);
    $ext->setAttribute("cy", $height * 9525);
    $xfrm->appendChild($ext);
    $spPr->appendChild($xfrm);
    $prstGeom = $doc->createElementNS($ns_a, 'a:prstGeom');
    $prstGeom->setAttribute("prst", "rect");
    $avLst = $doc->createElementNS($ns_a, 'a:avLst');
    $prstGeom->appendChild($avLst);
    $spPr->appendChild($prstGeom);
    $pic->appendChild($spPr);

    // Replace the placeholder node
    $parentRun = $refNode->parentNode;
    $parentRun->parentNode->replaceChild($r, $parentRun);
}

// -------------------- USAGE --------------------

$template = "input.docx";
$output = "output.docx";

$replacements = [
    "%*employerName*%" => "ACME Corp",
    "%*CaseID*%"       => "CASE12345",
    "%*cName*%"        => "John Doe",
    "%*reportStatus*%" => "Clear",
    "%*dob*%"          => "01-01-1990",
    "%*reuestDate*%"   => "20-08-2025",
    "%*cID*%"          => "C123",
    "%*reportDate*%"   => "21-08-2025",
    "%*img1*%"         => [
        "image"  => 'C:\xampp\htdocs\test\pdfgenback\openxml\image1.jpg',
        "width"  => 300,
        "height" => 300
    ]
];

try {
    replaceDocxPlaceholders($template, $output, $replacements);
    echo "✅ DOCX generated: $output";
} catch (Exception $e) {
    echo "Error: " . $e->getMessage();
}
?>