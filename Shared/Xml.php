<?php
// namespace MyPhpWord\Shared;
class Xml
{
//only document
    public function document($path)
    {
        $document = '';
        $relations = self::relations($path);
        $tt = [];
        $xmlparser = xml_parser_create();
        $xmldata = file_get_contents($path . '/word/document.xml');

        // Parse XML data into an array
        xml_parse_into_struct($xmlparser, $xmldata, $values);
        xml_parser_free($xmlparser);

        $ppr = false;
        $dom = new DOMDocument();
        $dom->formatOutput = true;
        $style = '';
        $styletbl = '';
        $tags = [];
        $wr = false;
        $tcborders = false;
        $color = '';
        foreach ($values as $value) {
            // if ($value['tag'] == 'W:DOCUMENT') {
            //     if ($value['type'] == 'open') {
            //         $document .= '<DOCUMENT>';
            //     } else {
            //         $document .= '</DOCUMENT>';
            //     };
            // }
            // if ($value['tag'] == 'W:BODY') {
            //     if ($value['type'] == 'open') {
            //         $document .= '<BODY>';
            //     } else {
            //         $document .= '</BODY>';
            //     };
            // }
            if ($value['tag'] == 'W:P') {
                if ($value['type'] == 'open') {
                    $document .= '<P>';
                    $style = '';
                }
                // if ($value['type'] == 'complite') {
                //     $document .= '<P></P>';
                // }
                if ($value['type'] == 'close') {
                    $document .= '</P>';
                };
            }
            if ($value['tag'] == 'W:PPR') {
                if ($value['type'] == 'open') {
                    $ppr = true;
                } else {
                    $ppr = false;
                    // $style = '';
                };
            }
            //
            if ($ppr) {
                if ($value['tag'] == 'W:JC') {
                    $style = ' style="text-align: ' . $value['attributes']['W:VAL'] . ';"';
                    $document = substr($document, 0, strlen($document) - 1);
                    $document .= $style . '>';
                }
                if ($value['tag'] == 'W:SZ') {
                    $sz = round((float) $value['attributes']['W:VAL'] / 1.8, 0);
                    if ($style != '') {
                        $document = substr($document, 0, strlen($document) - 2);
                        $document .= 'font-size: ' . $sz . 'pt;">';
                    } else {
                        $document = substr($document, 0, strlen($document) - 1);
                        $document .= ' style="font-size: ' . $sz . 'pt;">';
                    }
                }
            }

            if ($value['tag'] == 'W:TBL') {
                if ($value['type'] == 'open') {
                    $document .= '<TABLE>';
                    $styletbl = '';
                } else {
                    $document .= '</TABLE>';
                };
            }
            if ($value['tag'] == 'W:TBLW') {
                // if ($value['type'] == 'open') {
                $w = round((float) $value['attributes']['W:W'] / 20, 2);
                if ($w != 0) {
                    $styletbl = ' style="width: ' . $w . 'pt; border-collapse: collapse;"';
                } else {
                    $styletbl = ' style="width: 100%; border-collapse: collapse;"';
                }
                $document = substr($document, 0, strlen($document) - 1);
                $document .= $styletbl . '>';
                $styletbl = '';
            }

            if ($value['tag'] == 'W:TBLGRID') {
                if ($value['type'] == 'open') {
                    $document .= '<COLGROUP>';
                } else {
                    $document .= '</COLGROUP>';
                    // $document .= '</TABLE>';
                };
            }
            if ($value['tag'] == 'W:GRIDCOL') {
                $w = round((float) $value['attributes']['W:W'] / 20, 2);
                $document .= '<COL style="width: ' . $w . 'pt;">';
            }

            if ($value['tag'] == 'W:TR') {
                if ($value['type'] == 'open') {
                    $document .= '<TR>';
                } else {
                    $document .= '</TR>';
                };
            }

            if ($value['tag'] == 'W:TRHEIGHT') {
                $document = substr($document, 0, strlen($document) - 1);
                $h = round((float) $value['attributes']['W:VAL'] / 20, 2);
                $document .= ' style="height: ' . $h . 'pt;">';
            }

            if ($value['tag'] == 'W:TC') {
                if ($value['type'] == 'open') {
                    $document .= '<TD>';
                } else {
                    $document .= '</TD>';
                };
            }
            if ($value['tag'] == 'W:GRIDSPAN') {
                $document = substr($document, 0, strlen($document) - 1);
                $document .= ' colspan="' . $value['attributes']['W:VAL'] . '">';
            }

            if ($value['tag'] == 'W:TCBORDERS') {
                if ($value['type'] == 'open') {
                    $tcborders = true;
                    $styletc = '';
                } else {
                    if ($styletc != '') {
                        $document = substr($document, 0, strlen($document) - 1);
                        $document .= ' style="' . $styletc . '">';
                    }
                    $tcborders = false;
                };
            }

            if ($tcborders) {
                if ($value['tag'] == 'W:TOP') {
                    if (array_key_exists('W:SZ', $value['attributes'])) {
                        $styletc .= 'border-top: ' . round((float) $value['attributes']['W:SZ'] / 20, 2) . 'pt solid;';
                    }
                }
                if ($value['tag'] == 'W:RIGHT') {
                    if (array_key_exists('W:SZ', $value['attributes'])) {
                        $styletc .= 'border-right: ' . round((float) $value['attributes']['W:SZ'] / 20, 2) . 'pt solid;';
                    }
                }
                if ($value['tag'] == 'W:BOTTOM') {
                    if (array_key_exists('W:SZ', $value['attributes'])) {
                        $styletc .= 'border-bottom: ' . round((float) $value['attributes']['W:SZ'] / 20, 2) . 'pt solid;';
                    }
                }
                if ($value['tag'] == 'W:LEFT') {
                    if (array_key_exists('W:SZ', $value['attributes'])) {
                        $styletc .= 'border-left: ' . round((float) $value['attributes']['W:SZ'] / 20, 2) . 'pt solid;';
                    }
                }
            }
            if ($value['tag'] == 'W:R') {
                if ($value['type'] == 'open') {
                    $wr = true;
                } else {
                    $wr = false;
                    $tagend = '';
                    $color = '';
                    for ($i = count($tags); $i > 0; $i--) {
                        $tagend .= '</' . $tags[$i - 1] . '>';
                    }

                    $document .= $tagend;
                    $tags = [];
                };
            }
            //
            if ($value['tag'] == 'W:B') {
                if ($wr) {
                    $tags[] = 'STRONG';
                }
            }
            if ($value['tag'] == 'W:I') {
                if ($wr) {
                    $tags[] = 'I';
                }
            }
            if ($value['tag'] == 'W:U') {
                if ($wr) {
                    $tags[] = 'U';
                }
            }
            if ($value['tag'] == 'W:COLOR') {
                if ($wr) {
                    $tags[] = 'FONT';
                    $color = $value['attributes']['W:VAL'];
                }
            }

            //  if ($value['tag'] == 'W:TAB') {
            // if ($value['type'] == 'complete') {
            //      $document .= '&nbsp;&nbsp;';
            // }
            //  }

            if ($value['tag'] == 'W:T') {
                if (array_key_exists('value', $value)) {
                    $tagstart = '';
                    for ($i = 0; $i < count($tags); $i++) {
                        if ($tags[$i] == 'FONT') {
                            $tagstart .= '<' . $tags[$i] . ' color="' . $color . '">';
                        } else {
                            $tagstart .= '<' . $tags[$i] . '>';
                        }
                    }
                    $document .= $tagstart . $value['value'];
                }
                // $tt[]=$value;
            }
//IMAGE
            if ($value['tag'] == 'A:BLIP') {
                if ($value['type'] == 'open') {
                    $image = $path . '/word/' . $relations[$value['attributes']['R:EMBED']];
                    $document .= '<IMAGE src="' . $image . '" style="max-width:60%">';
                }
            }

//PAGE SIZE
            if ($value['tag'] == 'W:PGSZ') {
                $w = round((float) $value['attributes']['W:W'] / 20, 2);
                $h = round((float) $value['attributes']['W:H'] / 20, 2);
                $style = 'style="width: ' . $w . 'pt; min-height: ' . $h . 'pt;"';

            }
            if ($value['tag'] == 'W:PGMAR') {
                $top = round((float) $value['attributes']['W:TOP'] / 20, 2);
                $right = round((float) $value['attributes']['W:RIGHT'] / 20, 2);
                $bottom = round((float) $value['attributes']['W:BOTTOM'] / 20, 2);
                $left = round((float) $value['attributes']['W:LEFT'] / 20, 2);
                $style = 'style="padding: ' . $top . 'pt ' . $right . 'pt ' . $bottom . 'pt ' . $left . 'pt; width: ' . $w . 'pt; min-height: ' . $h . 'pt;"';
            }
        }

        $document = preg_replace('/<\/STRONG><STRONG>/', '', $document);
        $document = preg_replace('/<\/I><I>/', '', $document);
        $document = preg_replace('/<\/U><U>/', '', $document);
        $document = preg_replace('/<STRONG> <\/STRONG>/', ' ', $document);
        $document = preg_replace('/<STRONG> /', ' <STRONG>', $document);
        $document = preg_replace('/<\/FONT><FONT.*?>/', '', $document);

        $document = '<style>.docx-wrapper {
    background: gray;
    padding: 30px;
    padding-bottom: 0px;
    display: flex;
    flex-flow: column;
    align-items: center;
}
.docx-wrapper>section.docx {
    background: white;
    box-shadow: 0 0 10px rgba(0, 0, 0, 0.5);
    margin-bottom: 30px;
}
    @media print {
    .docx-wrapper { background: none; padding: none; padding-bottom: none; display: block; }
    .docx-wrapper>section.docx { background: white; box-shadow: none; margin-bottom: none;}
    section.docx { page-break-after: always; }
    section.docx>article { margin-bottom: none; }
}
    </style><div class="docx-wrapper"><section class="docx" ' . $style . '>' . $document;
        $document .= '</section></div>';

        return $document;
        // return json_encode($values);
        // return json_encode($tt);

    }
    private function relations($path)
    {
        $relations = [];
        $xmlparser = xml_parser_create();
        $xmldata = file_get_contents($path . '/word/_rels/document.xml.rels');

        // Parse XML data into an array
        xml_parse_into_struct($xmlparser, $xmldata, $values);
        xml_parser_free($xmlparser);
        foreach ($values as $value) {
            if ($value['tag'] == 'RELATIONSHIP') {
                $relations[$value['attributes']['ID']] = $value['attributes']['TARGET'];
            }
        }
        $xmlparser = null;
        return $relations;
    }

//     <w:document
// xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
// xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
// xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
// xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
// xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"
// xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
// xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"
// xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"
// xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"
// xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"
// xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
// xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink"
// xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"
// xmlns:o="urn:schemas-microsoft-com:office:office"
// xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
// xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
// xmlns:v="urn:schemas-microsoft-com:vml"
// xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
// xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
// xmlns:w10="urn:schemas-microsoft-com:office:word"
// xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
// xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
// xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
// xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
// xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
// xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
// xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
// xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
// xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
// mc:Ignorable="w14 w15 w16se w16cid wp14">

}
