"""
Add ElevenLabs TTS audio narration to a PowerPoint file.

Requirements:
    pip install python-pptx requests

Usage:
    python add_audio_to_pptx.py

The script reads speaker notes from each slide, generates MP3 audio via
ElevenLabs, and embeds each audio file into the corresponding slide so it
plays automatically when the slide is shown.
"""

import io
import os
import time
import requests
from lxml import etree
from pptx import Presentation
from pptx.util import Pt

# ── Configuration ────────────────────────────────────────────────────────────
ELEVENLABS_API_KEY = "3edf18bfe491b6d56480eb4b801091d3b735823f28858f8d9080194f6abed971"
VOICE_ID = "JBFqnCBsd6RMkjVDRZzb"  # "George" — change to any ElevenLabs voice ID
MODEL_ID = "eleven_multilingual_v2"

INPUT_PPTX  = "training_week_6_with_narration.pptx"   # path to your input file
OUTPUT_PPTX = "training_week_6_with_audio.pptx"       # path for the output file
# ─────────────────────────────────────────────────────────────────────────────


def get_slide_notes(slide) -> str:
    if slide.has_notes_slide:
        return slide.notes_slide.notes_text_frame.text.strip()
    return ""


def generate_audio(text: str, api_key: str, voice_id: str, model_id: str) -> bytes:
    """Call ElevenLabs TTS API and return MP3 bytes."""
    url = f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}"
    headers = {
        "xi-api-key": api_key,
        "Content-Type": "application/json",
        "Accept": "audio/mpeg",
    }
    payload = {
        "text": text,
        "model_id": model_id,
        "voice_settings": {
            "stability": 0.5,
            "similarity_boost": 0.75,
        },
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    return resp.content


def embed_audio_in_slide(prs: Presentation, slide, audio_bytes: bytes, slide_index: int):
    """
    Embed an MP3 into a slide as a hidden media element that auto-plays.
    Uses the OOXML relationship + timing model supported by PowerPoint.
    """
    # Add the audio file to the presentation package
    audio_stream = io.BytesIO(audio_bytes)
    media_rId = slide.part.relate_to(
        audio_stream,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio",
        is_external=False,
    )

    # Build the <p:sp> picture element that holds the audio icon (hidden, 1×1 px)
    spTree = slide.shapes._spTree
    nsmap = {
        "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
        "p":   "http://schemas.openxmlformats.org/presentationml/2006/main",
        "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    }

    pic_xml = f"""<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvPicPr>
    <p:cNvPr id="{100 + slide_index}" name="Audio {slide_index + 1}"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr>
      <p:ph type="body" idx="1"/>
      <p:audioFile r:link="{media_rId}"/>
    </p:nvPr>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>"""

    pic_el = etree.fromstring(pic_xml)
    spTree.append(pic_el)

    # Add timing so audio plays automatically on slide entry
    timing_xml = f"""<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="whenNotActive" nodeType="tmRoot">
        <p:childTnLst>
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst>
                <p:par>
                  <p:cTn id="3" fill="hold">
                    <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
                    <p:childTnLst>
                      <p:par>
                        <p:cTn id="4" fill="hold">
                          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                          <p:childTnLst>
                            <p:audio>
                              <p:cMediaNode vol="80000" mute="0" showWhenStopped="0" r:id="{media_rId}"
                                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                                <p:cTn id="5" fill="hold">
                                  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                </p:cTn>
                                <p:tgtEl>
                                  <p:spTgt spid="{100 + slide_index}"/>
                                </p:tgtEl>
                              </p:cMediaNode>
                            </p:audio>
                          </p:childTnLst>
                        </p:cTn>
                      </p:par>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
            <p:prevCondLst>
              <p:cond evt="onPrev" delay="0">
                <p:tgtEl><p:sldTgt/></p:tgtEl>
              </p:cond>
            </p:prevCondLst>
            <p:nextCondLst>
              <p:cond evt="onNext" delay="0">
                <p:tgtEl><p:sldTgt/></p:tgtEl>
              </p:cond>
            </p:nextCondLst>
          </p:seq>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
  <p:bldLst/>
</p:timing>"""

    # Replace or append timing element on slide XML
    slide_el = slide._element
    existing_timing = slide_el.find(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}timing"
    )
    if existing_timing is not None:
        slide_el.remove(existing_timing)
    slide_el.append(etree.fromstring(timing_xml))


def main():
    print(f"Loading: {INPUT_PPTX}")
    prs = Presentation(INPUT_PPTX)
    total = len(prs.slides)

    for i, slide in enumerate(prs.slides):
        notes = get_slide_notes(slide)
        if not notes:
            print(f"  Slide {i+1}/{total}: no notes, skipping")
            continue

        print(f"  Slide {i+1}/{total}: generating audio ({len(notes)} chars)...", end=" ", flush=True)
        try:
            audio = generate_audio(notes, ELEVENLABS_API_KEY, VOICE_ID, MODEL_ID)
            embed_audio_in_slide(prs, slide, audio, i)
            print(f"done ({len(audio)//1024} KB)")
        except Exception as e:
            print(f"ERROR: {e}")

        # Brief pause to stay within ElevenLabs rate limits
        time.sleep(0.5)

    prs.save(OUTPUT_PPTX)
    print(f"\nSaved: {OUTPUT_PPTX}")


if __name__ == "__main__":
    main()
