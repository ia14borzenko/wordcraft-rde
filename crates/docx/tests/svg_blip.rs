//! A picture read from a file can carry extensions in its `a:blip` extension list — Word 2016+
//! keeps a vector source there as `asvg:svgBlip` (MS-ODRAWXML) next to the raster preview. Saving
//! must write the extension (and the SVG part it names) back, not only the raster picture
//! (RDE P2.5, ADR-0005). Hand-written fixture from the spec: no files produced by Word.

use std::collections::BTreeMap;
use std::io::{Read, Write};

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const IMAGE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
/// The `a:ext` URI Word puts an SVG's `asvg:svgBlip` under (MS-ODRAWXML §2.3.1).
const SVG_BLIP_EXT: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const SVG_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

fn png() -> Vec<u8> {
    b"\x89PNG\r\n\x1a\n a raster preview of the vector picture".to_vec()
}

fn svg() -> Vec<u8> {
    br##"<?xml version="1.0" encoding="UTF-8"?><svg xmlns="http://www.w3.org/2000/svg" width="144" height="72"><rect width="144" height="72" fill="#156082"/></svg>"##.to_vec()
}

fn zip(entries: &[(&str, Vec<u8>)]) -> Vec<u8> {
    let mut zw = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for (name, bytes) in entries {
        zw.start_file(*name, zip::write::SimpleFileOptions::default()).unwrap();
        zw.write_all(bytes).unwrap();
    }
    zw.finish().unwrap().into_inner()
}

fn unzip(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let mut z = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut out = BTreeMap::new();
    for i in 0..z.len() {
        let mut f = z.by_index(i).unwrap();
        let mut b = Vec::new();
        f.read_to_end(&mut b).unwrap();
        out.insert(f.name().to_string(), b);
    }
    out
}

fn text(files: &BTreeMap<String, Vec<u8>>, name: &str) -> String {
    String::from_utf8(files.get(name).unwrap_or_else(|| panic!("no {name} in {:?}", files.keys())).clone()).unwrap()
}

fn rels(list: &[(&str, &str)]) -> Vec<u8> {
    let mut s = String::from(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns=""#);
    s.push_str(PKG_REL_NS);
    s.push_str(r#"">"#);
    for (id, target) in list {
        s.push_str(&format!(r#"<Relationship Id="{id}" Type="{IMAGE_REL}" Target="{target}"/>"#));
    }
    s.push_str("</Relationships>");
    s.into_bytes()
}

/// The picture's `w:drawing`: its `a:blip` keeps an SVG source (and an unrelated Office
/// extension) in its `a:extLst`.
fn drawing() -> String {
    format!(
        r#"<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"><a:extLst><a:ext uri="{SVG_BLIP_EXT}"><asvg:svgBlip xmlns:asvg="{SVG_NS}" r:embed="rIdSvg"/></a:ext><a:ext uri="{{28A0092B-C50C-407E-A947-70E740481C1C}}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#
    )
}

/// A package with one SVG-backed picture: `word/media/image1.png` (the preview) and
/// `word/media/image1.svg` (its vector source), related from the main document.
fn package() -> Vec<u8> {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{W_NS}" xmlns:r="{REL_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>{}<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#,
        drawing()
    );
    let root = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{PKG_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#
    );
    let doc_rels = rels(&[("rIdImg", "media/image1.png"), ("rIdSvg", "media/image1.svg")]);
    let types = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="{CT_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="svg" ContentType="image/svg+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#
    );
    zip(&[
        ("[Content_Types].xml", types.into_bytes()),
        ("_rels/.rels", root.into_bytes()),
        ("word/document.xml", document.into_bytes()),
        ("word/_rels/document.xml.rels", doc_rels),
        ("word/media/image1.png", png()),
        ("word/media/image1.svg", svg()),
    ])
}

/// Every `r:embed` the main document uses is related in it.
fn embeds_resolve(document: &str, rels: &str) -> bool {
    document.match_indices("r:embed=\"").all(|(i, m)| {
        let rest = &document[i + m.len()..];
        rest.split('"').next().is_some_and(|id| rels.contains(&format!(r#"Id="{id}""#)))
    })
}

#[test]
fn svg_blip_and_its_part_survive_open_and_save() {
    let doc = wordcraft_docx::read(&package()).expect("read");
    let image = doc
        .body
        .iter()
        .filter_map(|b| b.as_para())
        .flat_map(|p| p.objects.iter())
        .find_map(|o| match o {
            wordcraft_doc::InlineObject::Image { media, blip_ext, .. } => Some((media.clone(), blip_ext.clone())),
            _ => None,
        })
        .expect("the picture is read");
    assert!(image.0.ends_with(".png"), "the raster preview is the picture: {}", image.0);
    let ext = image.1.expect("the blip's extension list is kept");
    assert!(ext.xml.contains("svgBlip"), "the SVG extension is kept: {}", ext.xml);
    assert!(ext.xml.contains("useLocalDpi"), "the rest of the list is kept too: {}", ext.xml);
    let svg_ref = ext.refs.iter().find(|r| r.target.ends_with(".svg")).expect("the SVG part is named");
    assert!(!svg_ref.external);
    assert!(ext.refs.iter().any(|r| r.id == svg_ref.id && r.kind.ends_with("relationships/image")));
    assert_eq!(doc.passthrough.get("word/media/image1.svg").map(|b| b.as_slice()), Some(svg().as_slice()), "the SVG part itself is carried");

    let saved = unzip(&wordcraft_docx::write(&doc).expect("write"));
    let document = text(&saved, "word/document.xml");
    assert!(document.contains("svgBlip"), "svgBlip is written back: {document}");
    assert!(document.contains(SVG_BLIP_EXT), "under the extension's own URI");
    assert!(document.contains("useLocalDpi"), "the whole list is written back");
    assert_eq!(saved.get("word/media/image1.svg").map(|b| b.as_slice()), Some(svg().as_slice()), "the SVG part is written back");
    let rels = text(&saved, "word/_rels/document.xml.rels");
    assert!(rels.contains("media/image1.svg") && rels.contains("/relationships/image"), "the SVG part is related: {rels}");
    assert!(embeds_resolve(&document, &rels), "the written blip's relationships resolve: {document} / {rels}");
    let types = text(&saved, "[Content_Types].xml");
    assert!(types.contains("image/svg+xml"), "the SVG part's content type is written: {types}");

    // A second open/save keeps them too (the extension is not a one-shot read).
    let again = wordcraft_docx::read(&wordcraft_docx::write(&doc).unwrap()).unwrap();
    let saved = unzip(&wordcraft_docx::write(&again).unwrap());
    assert!(text(&saved, "word/document.xml").contains("svgBlip"));
    assert_eq!(saved.get("word/media/image1.svg").map(|b| b.as_slice()), Some(svg().as_slice()));
}
