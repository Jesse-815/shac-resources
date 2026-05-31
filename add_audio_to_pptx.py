"""
Add ElevenLabs TTS audio narration to a PowerPoint file.

Requirements:
    pip install requests lxml

Usage:
    python add_audio_to_pptx.py
"""

import os
import re
import time
import shutil
import zipfile
import requests
from lxml import etree

# ── Configuration ────────────────────────────────────────────────────────────
ELEVENLABS_API_KEY = "3edf18bfe491b6d56480eb4b801091d3b735823f28858f8d9080194f6abed971"
VOICE_ID = "JBFqnCBsd6RMkjVDRZzb"
MODEL_ID = "eleven_multilingual_v2"

INPUT_PPTX  = "training_week_6_with_narration (2).pptx"
OUTPUT_PPTX = "training_week_6_with_audio.pptx"
AUDIO_DIR   = "slide_audio"
WORK_DIR    = "_pptx_work"
# ─────────────────────────────────────────────────────────────────────────────

PPTNS = "http://schemas.openxmlformats.org/presentationml/2006/main"
ANS   = "http://schemas.openxmlformats.org/drawingml/2006/main"
RNS   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
AUDIO_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"


def generate_audio(text: str) -> bytes:
    url = f"https://api.elevenlabs.io/v1/text-to-speech/{VOICE_ID}"
    headers = {
        "xi-api-key": ELEVENLABS_API_KEY,
        "Content-Type": "application/json",
        "Accept": "audio/mpeg",
    }
    payload = {
        "text": text,
        "model_id": MODEL_ID,
        "voice_settings": {"stability": 0.5, "similarity_boost": 0.75},
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    return resp.content


def get_notes(slide_xml_path: str) -> str:
    """Extract speaker notes text from a slide XML file."""
    # Notes are in a separate file: ppt/notesSlides/notesSlideN.xml
    # Find the notes relationship from slide rels
    return ""  # handled below via zipfile


def get_notes_from_zip(zf: zipfile.ZipFile, slide_rel_path: str, slide_path: str) -> str:
    """Get notes text by following slide relationships."""
    try:
        rels_xml = zf.read(slide_rel_path)
        rels = etree.fromstring(rels_xml)
        for rel in rels.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
            if "notesSlide" in rel.get("Type", ""):
                target = rel.get("Target")
                # Resolve relative path
                slide_dir = "/".join(slide_path.split("/")[:-1])
                notes_path = slide_dir + "/" + target.lstrip("../").lstrip("./")
                # Handle ../ properly
                parts = (slide_dir + "/" + target).replace("\\", "/").split("/")
                resolved = []
                for p in parts:
                    if p == "..":
                        if resolved:
                            resolved.pop()
                    elif p and p != ".":
                        resolved.append(p)
                notes_path = "/".join(resolved)
                try:
                    notes_xml = zf.read(notes_path)
                    root = etree.fromstring(notes_xml)
                    # Get all text from txBody that isn't the slide number placeholder
                    texts = []
                    for sp in root.iter("{http://schemas.openxmlformats.org/presentationml/2006/main}sp"):
                        ph = sp.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}ph")
                        if ph is not None and ph.get("type") == "sldNum":
                            continue
                        for t in sp.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t"):
                            if t.text:
                                texts.append(t.text)
                    return "".join(texts).strip()
                except Exception:
                    pass
    except Exception:
        pass
    return ""


def add_audio_to_slide(zf_in: zipfile.ZipFile, zf_out: zipfile.ZipFile,
                       slide_path: str, audio_bytes: bytes, slide_index: int):
    """Copy slide XML with audio embedded, plus write the MP3 into the zip."""
    shape_id = 200 + slide_index
    media_name = f"audio{slide_index + 1}.mp3"
    media_path = f"ppt/media/{media_name}"

    # Write audio file into output zip
    zf_out.writestr(media_path, audio_bytes)

    # Update slide relationships
    slide_dir = "/".join(slide_path.split("/")[:-1])
    slide_name = slide_path.split("/")[-1].replace(".xml", "")
    rels_path = f"{slide_dir}/_rels/{slide_name}.xml.rels"

    try:
        rels_xml = zf_in.read(rels_path)
        rels_root = etree.fromstring(rels_xml)
    except Exception:
        rels_root = etree.fromstring(
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
        )

    # Find next available rId
    existing_ids = [r.get("Id", "") for r in rels_root]
    n = 1
    while f"rId{n}" in existing_ids:
        n += 1
    audio_rId = f"rId{n}"

    etree.SubElement(rels_root, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship", {
        "Id": audio_rId,
        "Type": AUDIO_REL_TYPE,
        "Target": f"../media/{media_name}",
    })
    zf_out.writestr(rels_path, etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True))

    # Update slide XML to add hidden audio shape + timing
    slide_xml = zf_in.read(slide_path)
    slide_root = etree.fromstring(slide_xml)

    csp_tree = slide_root.find(f"{{{PPTNS}}}cSld/{{{PPTNS}}}spTree")
    if csp_tree is None:
        # Try without namespace
        csp_tree = slide_root.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}spTree")

    pic_xml = (
        f'<p:pic xmlns:p="{PPTNS}" xmlns:a="{ANS}" xmlns:r="{RNS}">'
        f'<p:nvPicPr>'
        f'<p:cNvPr id="{shape_id}" name="Audio{slide_index+1}">'
        f'<a:hlinkClick r:id="" action="ppaction://media"/>'
        f'</p:cNvPr>'
        f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
        f'<p:nvPr><p:audioFile r:link="{audio_rId}"/></p:nvPr>'
        f'</p:nvPicPr>'
        f'<p:blipFill><a:blip r:embed="{audio_rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="457200" y="457200"/><a:ext cx="457200" cy="457200"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</p:spPr>'
        f'</p:pic>'
    )
    if csp_tree is not None:
        csp_tree.append(etree.fromstring(pic_xml))

    # Remove existing timing if any
    for timing in slide_root.findall(f"{{{PPTNS}}}timing"):
        slide_root.remove(timing)

    timing_xml = (
        f'<p:timing xmlns:p="{PPTNS}" xmlns:a="{ANS}">'
        f'<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="whenNotActive" nodeType="tmRoot">'
        f'<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
        f'<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par>'
        f'<p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
        f'<p:childTnLst><p:par><p:cTn id="4" fill="hold">'
        f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        f'<p:childTnLst><p:audio>'
        f'<p:cMediaNode vol="80000" mute="0" showWhenStopped="0">'
        f'<p:cTn id="5" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        f'<p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl>'
        f'</p:cMediaNode></p:audio></p:childTnLst>'
        f'</p:cTn></p:par></p:childTnLst></p:cTn>'
        f'</p:par></p:childTnLst></p:cTn>'
        f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        f'</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst/></p:timing>'
    )
    slide_root.append(etree.fromstring(timing_xml))

    zf_out.writestr(slide_path, etree.tostring(slide_root, xml_declaration=True, encoding="UTF-8", standalone=True))


def main():
    os.makedirs(AUDIO_DIR, exist_ok=True)

    print(f"Loading: {INPUT_PPTX}")
    with zipfile.ZipFile(INPUT_PPTX, "r") as zf:
        all_files = zf.namelist()

    # Find slide files in order
    slide_files = sorted(
        [f for f in all_files if re.match(r"ppt/slides/slide\d+\.xml$", f)],
        key=lambda x: int(re.search(r"\d+", x.split("/")[-1]).group())
    )
    total = len(slide_files)
    print(f"Found {total} slides")

    # Gather audio for each slide
    audio_map = {}  # slide_path -> audio_bytes
    with zipfile.ZipFile(INPUT_PPTX, "r") as zf:
        for i, slide_path in enumerate(slide_files):
            slide_name = slide_path.split("/")[-1].replace(".xml", "")
            rels_path = f"ppt/slides/_rels/{slide_name}.xml.rels"
            notes = get_notes_from_zip(zf, rels_path, slide_path)

            if not notes:
                print(f"  Slide {i+1}/{total}: no notes, skipping")
                continue

            audio_file = os.path.join(AUDIO_DIR, f"slide_{i+1:02d}.mp3")
            if os.path.exists(audio_file):
                print(f"  Slide {i+1}/{total}: using cached audio", end=" ", flush=True)
                with open(audio_file, "rb") as f:
                    audio = f.read()
            else:
                print(f"  Slide {i+1}/{total}: generating audio ({len(notes)} chars)...", end=" ", flush=True)
                try:
                    audio = generate_audio(notes)
                    with open(audio_file, "wb") as f:
                        f.write(audio)
                except Exception as e:
                    print(f"ERROR: {e}")
                    continue
                time.sleep(0.5)

            print(f"done ({len(audio)//1024} KB)")
            audio_map[slide_path] = audio

    # Build output PPTX
    print(f"\nBuilding output file...")
    with zipfile.ZipFile(INPUT_PPTX, "r") as zf_in:
        with zipfile.ZipFile(OUTPUT_PPTX, "w", zipfile.ZIP_DEFLATED) as zf_out:
            slides_with_audio = set(audio_map.keys())
            slides_rels_written = set()

            for item in zf_in.namelist():
                # Fix [Content_Types].xml to include mp3
                if item == "[Content_Types].xml":
                    ct_xml = zf_in.read(item)
                    ct_root = etree.fromstring(ct_xml)
                    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
                    existing_exts = {el.get("Extension") for el in ct_root.findall(f"{{{ct_ns}}}Default")}
                    if "mp3" not in existing_exts:
                        etree.SubElement(ct_root, f"{{{ct_ns}}}Default", {
                            "Extension": "mp3",
                            "ContentType": "audio/mpeg",
                        })
                    zf_out.writestr(item, etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True))
                    continue
                # Skip slide files and their rels that we'll rewrite
                slide_name_match = re.match(r"ppt/slides/(slide\d+)\.xml$", item)
                rels_match = re.match(r"ppt/slides/_rels/(slide\d+)\.xml\.rels$", item)

                if slide_name_match:
                    slide_path = item
                    if slide_path in slides_with_audio:
                        # Will be written by add_audio_to_slide
                        continue
                    else:
                        zf_out.writestr(item, zf_in.read(item))
                elif rels_match:
                    slide_path = f"ppt/slides/{rels_match.group(1)}.xml"
                    if slide_path in slides_with_audio:
                        # Will be written by add_audio_to_slide
                        continue
                    else:
                        zf_out.writestr(item, zf_in.read(item))
                else:
                    zf_out.writestr(item, zf_in.read(item))

            # Now write slides with audio
            with zipfile.ZipFile(INPUT_PPTX, "r") as zf_in2:
                for i, slide_path in enumerate(slide_files):
                    if slide_path in slides_with_audio:
                        add_audio_to_slide(zf_in2, zf_out, slide_path, audio_map[slide_path], i)

    print(f"Saved: {OUTPUT_PPTX}")


if __name__ == "__main__":
    main()
