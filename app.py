import streamlit as st
import anthropic
import json
import base64
import os
from dotenv import load_dotenv
from io import BytesIO
import re

from pptx import Presentation
from pptx.util import Inches
from pptx.chart.data import ChartData, CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_DATA_LABEL_POSITION

# ------------- 1) SETUP -------------
load_dotenv()
st.set_page_config(page_title="Flexible Slides from PDF", layout="wide")

API_KEY = os.getenv("ANTHROPIC_API_KEY")
if not API_KEY:
    st.error("Anthropic API key not found. Please set it in .env or environment variables.")
    st.stop()

client = anthropic.Anthropic(
    api_key=API_KEY,
    default_headers={"anthropic-beta": "pdfs-2024-09-25"}
)
MODEL_NAME = "claude-3-5-sonnet-20241022"  # or your chosen model

EXETER_TEMPLATE_PATH = "University-of-Exeter_Powerpoint_templates_16-9.pptx"
PROMPT_FILE = "prompt.txt"  # <-- Path to your .txt file with placeholders

# ------------- 2) UTILITY: Load the Prompt Template -------------
def load_prompt_text(prompt_path):
    """Load the entire prompt from a text file (with placeholders like {{FICATION_LEVEL}}, etc.)."""
    with open(prompt_path, "r", encoding="utf-8") as f:
        return f.read()

def get_slide_layouts(pptx_path):
    """
    Retrieves slide layout information including layout names and placeholder indices.
    Returns a list of dictionaries with 'layout_name' and 'placeholders'.
    """
    prs = Presentation(pptx_path)
    layout_info = []
    slide_master = prs.slide_masters[0]
    for layout in slide_master.slide_layouts:
        placeholders = []
        for shape in layout.placeholders:
            placeholders.append({
                "placeholder_idx": shape.placeholder_format.idx,
                "placeholder_name": shape.name,
                "shape_type": str(shape.shape_type)
            })
        layout_info.append({
            "layout_name": layout.name,
            "placeholders": placeholders
        })
    return layout_info

# ------------- 3) BUILD PROMPT -------------
def build_flexible_prompt(layout_info, fication_level, image_filenames):
    """
    Reads the .txt prompt and replaces {{FICATION_LEVEL}}, {{LAYOUT_INFO_JSON}}, and {{AVAILABLE_IMAGES}}.
    """
    prompt_template = load_prompt_text(PROMPT_FILE)

    # Convert layout info to JSON
    layout_info_json = json.dumps(layout_info, indent=2)

    # Convert image filenames to a JSON array (or a comma-separated list)
    # We'll do a simple JSON array:
    image_filenames_json = json.dumps(image_filenames, indent=2)

    # Replace placeholders in the prompt
    prompt_filled = (
        prompt_template
        .replace("{{FICATION_LEVEL}}", str(fication_level))
        .replace("{{LAYOUT_INFO_JSON}}", layout_info_json)
        .replace("{{AVAILABLE_IMAGES}}", image_filenames_json)
    )
    print(prompt_filled)
    return prompt_filled

# ------------- 4) CALL CLAUDE -------------
def call_claude_for_slides(pdf_bytes, layout_info, fication_level, image_filenames):
    """
    Calls Claude with the PDF, layout info, and available images to get slides JSON.
    """
    # Build your prompt string
    final_prompt = build_flexible_prompt(layout_info, fication_level, image_filenames)

    # Encode the PDF to base64
    pdf_b64 = base64.b64encode(pdf_bytes).decode("utf-8")

    # Build the messages array
    messages = [
        {
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": pdf_b64
                    }
                },
                {
                    "type": "text",
                    "text": final_prompt
                }
            ]
        }
    ]

    # Call Claude
    response = client.messages.create(
        model=MODEL_NAME,
        messages=messages,
        max_tokens=8192
    )

    assistant_reply = response.content[0].text

    # Attempt to extract JSON from <json_output> ... </json_output>
    json_pattern = r"<json_output>\s*(\{.*?\})\s*</json_output>"
    match = re.search(json_pattern, assistant_reply, re.DOTALL)
    if match:
        try:
            raw_json = match.group(1)
            return json.loads(raw_json)
        except json.JSONDecodeError:
            st.error("Claude's <json_output> isn't valid JSON. Full reply:")
            st.write(assistant_reply)
            return None
    else:
        # fallback: try entire reply as JSON
        try:
            return json.loads(assistant_reply)
        except:
            st.error("No valid JSON found in Claude's reply. Full reply:")
            st.write(assistant_reply)
            return None

# ------------- 5) CREATE PPT FROM JSON -------------
def find_placeholder_by_idx(slide, idx):
    """Returns the shape on `slide` whose placeholder_format.idx == idx, or None."""
    for shape in slide.placeholders:
        if shape.placeholder_format.idx == idx:
            return shape
    return None

def create_slides_from_json(prs, slides_json, layout_info, uploaded_images=None):
    """
    Populate slides in the given Presentation `prs` based on `slides_json`.
    Each slide in slides_json["slides"]:
      {
        "layout_name": "<layout_name>",
        "placeholders": {
          "<placeholder_idx>": "Content" or { "image_key": "filename" } or { "chart_type": "...", "chart_data": ... },
          ...
        }
      },
      ...
    """
    slide_master = prs.slide_masters[0]
    
    # Create a mapping from layout names to layout objects
    # (Zipping layout_info with slide_master.slide_layouts in the same order)
    layout_name_map = {
        info["layout_name"]: layout_obj
        for info, layout_obj in zip(layout_info, slide_master.slide_layouts)
    }

    for slide_def in slides_json.get("slides", []):
        # 1) Pick the layout by name
        layout_name = slide_def.get("layout_name")
        if not layout_name:
            st.warning("Slide definition missing 'layout_name'. Skipping slide.")
            continue

        layout = layout_name_map.get(layout_name)
        if not layout:
            st.warning(f"Layout name '{layout_name}' not found in template. Skipping slide.")
            continue

        slide = prs.slides.add_slide(layout)

        # 2) Fill each placeholder
        placeholder_defs = slide_def.get("placeholders", {})
        for placeholder_idx, content in placeholder_defs.items():
            try:
                ph_idx = int(placeholder_idx)
            except ValueError:
                st.warning(f"Invalid placeholder index '{placeholder_idx}'. Must be an integer.")
                continue

            shape = find_placeholder_by_idx(slide, ph_idx)
            if not shape:
                st.warning(f"Placeholder with idx '{ph_idx}' not found in layout '{layout_name}'.")
                continue

            # (A) Handle dictionary-based content like images or charts
            if isinstance(content, dict):
                # Chart?
                if "chart_type" in content and "chart_data" in content:
                    chart_type = content["chart_type"]
                    chart_data = content["chart_data"]
                    # Remove placeholder shape
                    sp = shape._element
                    sp.getparent().remove(sp)
                    left, top, width, height = shape.left, shape.top, shape.width, shape.height
                    if chart_type == "donut":
                        _add_donut_chart(slide, left, top, width, height, chart_data)
                    elif chart_type == "comparison_bars":
                        _add_comparison_bars_chart(slide, left, top, width, height, chart_data)
                    elif chart_type == "trend_line":
                        _add_trend_line_chart(slide, left, top, width, height, chart_data)
                    continue

                # Image?
                if "image_key" in content and uploaded_images:
                    img_key = content["image_key"]
                    if img_key in uploaded_images:
                        # Remove placeholder shape
                        sp = shape._element
                        sp.getparent().remove(sp)
                        left, top, width, height = shape.left, shape.top, shape.width, shape.height
                        slide.shapes.add_picture(BytesIO(uploaded_images[img_key]), left, top, width, height)
                    continue

                # Possibly text + bullets
                text_val = content.get("text", "")
                bullet_vals = content.get("bullets", [])
                if text_val or bullet_vals:
                    tf = shape.text_frame
                    tf.text = text_val
                    for b_item in bullet_vals:
                        p = tf.add_paragraph()
                        p.text = b_item
                        p.level = 0

            # (B) If it's a simple string, treat as text
            elif isinstance(content, str):
                tf = shape.text_frame
                tf.text = content

    return prs

# -------------------- EXAMPLE CHART UTILS --------------------
def _add_donut_chart(slide, left, top, width, height, data_list):
    """
    Adds a donut chart to the slide.
    data_list: e.g. [ {"category": "Female", "value": 78.6}, {"category": "Male", "value": 21.4} ]
    """
    chart_data = ChartData()
    cats = [d["category"] for d in data_list]
    vals = [d["value"] for d in data_list]
    chart_data.categories = cats
    chart_data.add_series("Series 1", vals)

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT,
        left, top, width, height,
        chart_data
    ).chart
    chart.plots[0].donut_hole_size = 60
    chart.has_title = True
    chart.chart_title.text_frame.text = "Donut Chart"

def _add_comparison_bars_chart(slide, left, top, width, height, data_obj):
    """
    Adds a comparison bars chart to the slide.
    data_obj: e.g. {"labels": ["Group A","Group B"], "values": [75,45]}
    """
    chart_data = CategoryChartData()
    labels = data_obj.get("labels", [])
    values = data_obj.get("values", [])
    chart_data.categories = labels
    chart_data.add_series("Series 1", values)

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        left, top, width, height,
        chart_data
    ).chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "Bar Chart"

def _add_trend_line_chart(slide, left, top, width, height, data_obj):
    """
    Adds a trend line chart to the slide.
    data_obj: e.g. {"dates": ["Jan","Feb"], "values": [10,20]}
    """
    chart_data = CategoryChartData()
    dates = data_obj.get("dates", [])
    vals = data_obj.get("values", [])
    chart_data.categories = dates
    chart_data.add_series("Series 1", vals)

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        left, top, width, height,
        chart_data
    ).chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "Trend Line"

# ------------- 6) MAIN STREAMLIT APP -------------
def main():
    st.title("Let's Create and Edit Your Presentation")

    # Step A: Load layout info once
    if "layout_info" not in st.session_state:
        st.session_state.layout_info = get_slide_layouts(EXETER_TEMPLATE_PATH)

    # Step B: user picks fication level
    st.markdown("#### Choose simplification Level")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col1:
        st.markdown("**Academic**")
    with col2:
        simplification_level = st.slider("", 1, 10, 5)
    with col3:
        st.markdown("**Patient**")

    # Step C: user uploads PDF
    uploaded_pdf = st.file_uploader("Upload PDF", type=["pdf"])

    # Optional: user uploads images
    uploaded_images_files = st.file_uploader(
        "Upload Images", 
        type=["png", "jpg", "jpeg"], 
        accept_multiple_files=True
    )
    uploaded_images = {}
    image_filenames = []
    if uploaded_images_files:
        for img in uploaded_images_files:
            # We'll store the file by its actual name
            uploaded_images[img.name] = img.read()
            image_filenames.append(img.name)

    # Step D: call Claude to get slides JSON
    if uploaded_pdf and st.button("Let's Peel!"):
        with st.spinner("Generating slides..."):
            pdf_bytes = uploaded_pdf.read()
            result = call_claude_for_slides(
                pdf_bytes,
                st.session_state.layout_info,
                simplification_level,
                image_filenames  # pass the list of available image filenames
            )
            if result:
                st.session_state.slides_json = result
                st.success("Received JSON from Claude!")
            else:
                st.error("No valid JSON returned")

    # If we have slides_json, show editing interface
    if "slides_json" in st.session_state and st.session_state.slides_json:
        st.header("Edit Generated Slides")

        slides = st.session_state.slides_json.get("slides", [])

        # Slide selection
        st.subheader("Select a Slide to Edit")
        slide_options = [f"Slide {i+1}" for i in range(len(slides))]
        selected_slide_label = st.radio(
            "Choose a slide:",
            options=slide_options,
            index=0,
            key="selected_slide_radio"
        )
        selected_slide_index = slide_options.index(selected_slide_label)

        st.markdown(f"### Editing Slide {selected_slide_index + 1}")

        selected_slide = slides[selected_slide_index]
        placeholders = selected_slide.get("placeholders", {})

        for ph_idx, content in placeholders.items():
            # Find placeholder name
            placeholder_info = None
            for layout in st.session_state.layout_info:
                if layout['layout_name'] == selected_slide.get("layout_name", ""):
                    for ph in layout['placeholders']:
                        if str(ph['placeholder_idx']) == str(ph_idx):
                            placeholder_info = ph
                            break
            if not placeholder_info:
                st.warning(f"Placeholder idx {ph_idx} not found in layout info.")
                continue

            ph_name = placeholder_info['placeholder_name']
            display_name = "Title" if "title" in ph_name.lower() else ph_name

            st.markdown(f"**{display_name}**")

            # Check if content is dict with chart/image
            if isinstance(content, dict):
                if "chart_type" in content:
                    st.info("Chart placeholders are not editable via this interface.")
                    continue
                if "image_key" in content:
                    img_key = content["image_key"]
                    if img_key in uploaded_images:
                        st.image(uploaded_images.get(img_key, b''), caption=img_key, use_column_width=True)
                        st.info("Image placeholders are not editable via this interface.")
                    continue

                # Possibly text + bullets
                text = content.get("text", "")
                bullets = content.get("bullets", [])
            else:
                text = content
                bullets = []

            # Editable text
            edited_text = st.text_area(
                f"Edit Text for {display_name}",
                value=text,
                key=f"text_{selected_slide_index}_{ph_idx}",
                height=100
            )

            # Editable bullets
            edited_bullets = []
            if bullets:
                bullets_text = "\n".join(bullets)
                edited_bullets_text = st.text_area(
                    f"Edit Bullets for {display_name} (one per line)",
                    value=bullets_text,
                    key=f"bullets_{selected_slide_index}_{ph_idx}"
                )
                edited_bullets = [line.strip() for line in edited_bullets_text.split("\n") if line.strip()]

            # Update in session
            if bullets or edited_bullets:
                # If bullets exist
                if isinstance(content, dict):
                    slides[selected_slide_index]['placeholders'][ph_idx]['text'] = edited_text
                    slides[selected_slide_index]['placeholders'][ph_idx]['bullets'] = edited_bullets
                else:
                    slides[selected_slide_index]['placeholders'][ph_idx] = {
                        "text": edited_text,
                        "bullets": edited_bullets
                    }
            else:
                # No bullets
                if isinstance(content, dict):
                    slides[selected_slide_index]['placeholders'][ph_idx]['text'] = edited_text
                else:
                    slides[selected_slide_index]['placeholders'][ph_idx] = edited_text

        st.session_state.slides_json['slides'] = slides
        st.success("Slides JSON updated with your edits.")

        with st.expander("View Updated JSON"):
            st.json(st.session_state.slides_json)

        # Generate PPTX
        st.header("Generate Your Edited Presentation")
        if st.button("Generate PPT"):
            with st.spinner("Generating PowerPoint presentation..."):
                prs = Presentation(EXETER_TEMPLATE_PATH)
                prs = create_slides_from_json(prs, st.session_state.slides_json, st.session_state.layout_info, uploaded_images)
                ppt_buffer = BytesIO()
                prs.save(ppt_buffer)
                ppt_buffer.seek(0)
                st.download_button(
                    "Download PPTX",
                    data=ppt_buffer.getvalue(),
                    file_name="my_presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                st.success("PPTX generated and ready for download!")

if __name__ == "__main__":
    main()
