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
st.set_page_config(page_title="PEEL", layout="wide")

# Custom CSS for just sizing adjustments
st.markdown("""
    <style>
    .block-container {
        padding-top: 1rem;
        padding-bottom: 0rem;
    }
    div[data-testid="stFileUploader"] {
        width: 100%;
        padding: 1rem;
    }
    /* Custom styling for Let's Peel button */
    .stButton > button, .stDownloadButton > button {
        background-color: #e7fd7d !important;
        color: #544ff0 !important;
        font-size: 1.2em !important;
        padding: 0.8em 1.6em !important;
        transition: all 0.3s ease !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background-color: #d9fc5c !important;
        transform: scale(1.05);
    }
    </style>
""", unsafe_allow_html=True)

API_KEY = os.getenv("ANTHROPIC_API_KEY")
if not API_KEY:
    st.error("Anthropic API key not found. Please set it in .env or environment variables.")
    st.stop()

client = anthropic.Anthropic(
    api_key=API_KEY,
    default_headers={"anthropic-beta": "pdfs-2024-09-25"}
)
MODEL_NAME = "claude-3-5-sonnet-20241022"

EXETER_TEMPLATE_PATH = "exetertemplate.pptx"
PROMPT_FILE = "prompt.txt"

def load_prompt_text(prompt_path):
    with open(prompt_path, "r", encoding="utf-8") as f:
        return f.read()

def get_slide_layouts(pptx_path):
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

def build_flexible_prompt(layout_info, simplification_level, image_filenames):
    prompt_template = load_prompt_text(PROMPT_FILE)
    layout_info_json = json.dumps(layout_info, indent=2)
    image_filenames_json = json.dumps(image_filenames, indent=2)
    
    prompt_filled = (
        prompt_template
        .replace("{{SIMPLIFICATION_LEVEL}}", str(simplification_level))
        .replace("{{LAYOUT_INFO_JSON}}", layout_info_json)
        .replace("{{AVAILABLE_IMAGES}}", image_filenames_json)
    )
    return prompt_filled

def call_claude_for_slides(pdf_bytes, layout_info, simplification_level, image_filenames):
    final_prompt = build_flexible_prompt(layout_info, simplification_level, image_filenames)
    pdf_b64 = base64.b64encode(pdf_bytes).decode("utf-8")
    
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

    response = client.messages.create(
        model=MODEL_NAME,
        messages=messages,
        max_tokens=8192
    )

    assistant_reply = response.content[0].text
    json_pattern = r"<json_output>\s*(\{.*?\})\s*</json_output>"
    match = re.search(json_pattern, assistant_reply, re.DOTALL)
    
    if match:
        try:
            raw_json = match.group(1)
            return json.loads(raw_json)
        except json.JSONDecodeError:
            st.error("Invalid JSON in Claude's response.")
            return None
    else:
        try:
            return json.loads(assistant_reply)
        except:
            st.error("No valid JSON found in Claude's response.")
            return None

def find_placeholder_by_idx(slide, idx):
    for shape in slide.placeholders:
        if shape.placeholder_format.idx == idx:
            return shape
    return None

def create_slides_from_json(prs, slides_json, layout_info, uploaded_images=None):
    slide_master = prs.slide_masters[0]
    layout_name_map = {
        info["layout_name"]: layout_obj
        for info, layout_obj in zip(layout_info, slide_master.slide_layouts)
    }

    for slide_def in slides_json.get("slides", []):
        layout_name = slide_def.get("layout_name")
        if not layout_name:
            continue

        layout = layout_name_map.get(layout_name)
        if not layout:
            continue

        slide = prs.slides.add_slide(layout)
        placeholder_defs = slide_def.get("placeholders", {})
        
        for placeholder_idx, content in placeholder_defs.items():
            try:
                ph_idx = int(placeholder_idx)
            except ValueError:
                continue

            shape = find_placeholder_by_idx(slide, ph_idx)
            if not shape:
                continue

            if isinstance(content, dict):
                if "chart_type" in content and "chart_data" in content:
                    sp = shape._element
                    sp.getparent().remove(sp)
                    left, top, width, height = shape.left, shape.top, shape.width, shape.height
                    
                    chart_type = content["chart_type"]
                    chart_data = content["chart_data"]
                    
                    if chart_type == "donut":
                        _add_donut_chart(slide, left, top, width, height, chart_data)
                    elif chart_type == "comparison_bars":
                        _add_comparison_bars_chart(slide, left, top, width, height, chart_data)
                    elif chart_type == "trend_line":
                        _add_trend_line_chart(slide, left, top, width, height, chart_data)
                    continue

                if "image_key" in content and uploaded_images:
                    img_key = content["image_key"]
                    if img_key in uploaded_images:
                        sp = shape._element
                        sp.getparent().remove(sp)
                        left, top, width, height = shape.left, shape.top, shape.width, shape.height
                        slide.shapes.add_picture(BytesIO(uploaded_images[img_key]), left, top, width, height)
                    continue

                text_val = content.get("text", "")
                bullet_vals = content.get("bullets", [])
                if text_val or bullet_vals:
                    tf = shape.text_frame
                    tf.text = text_val
                    for b_item in bullet_vals:
                        p = tf.add_paragraph()
                        p.text = b_item
                        p.level = 0

            elif isinstance(content, str):
                tf = shape.text_frame
                tf.word_wrap = True
                tf.text = content

    return prs

def _add_donut_chart(slide, left, top, width, height, data_list):
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
    chart_data = CategoryChartData()
    chart_data.categories = data_obj.get("labels", [])
    chart_data.add_series("Series 1", data_obj.get("values", []))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        left, top, width, height,
        chart_data
    ).chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "Bar Chart"

def _add_trend_line_chart(slide, left, top, width, height, data_obj):
    chart_data = CategoryChartData()
    chart_data.categories = data_obj.get("dates", [])
    chart_data.add_series("Series 1", data_obj.get("values", []))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        left, top, width, height,
        chart_data
    ).chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "Trend Line"

def main():
    # Sidebar authentication
    with st.sidebar:
        st.title("Authentication")
        password = st.text_input("Enter Password", type="password")
        
        # Disclaimer text
        st.markdown("---")
        st.markdown("""
        **Disclaimer:**
        This tool uses AI to summarise medical documents, but it may contain errors or omissions. Always verify critical information with trusted medical sources and consult a qualified professional before making any decisions based on the content generated. The AI does not provide medical advice, diagnosis, or treatment. Use at your own discretion.


        """)

    # Check password
    if password != "P33L2025":
        st.error("Please enter the correct password to access the application.")
        return

    # Main application content starts here
    st.title("1. Create your presentation")

    if "layout_info" not in st.session_state:
        st.session_state.layout_info = get_slide_layouts(EXETER_TEMPLATE_PATH)

    st.header("Select simplification level")
    simplification_level = st.select_slider(
        "",
        options=list(range(1, 11)),
        format_func=lambda x: "Academic" if x == 1 else "Patient" if x == 10 else f"Level {x}",
        value=5
    )

    st.header("Upload pdf")
    uploaded_pdf = st.file_uploader("", type=["pdf"])

    st.header("Upload images")
    uploaded_images_files = st.file_uploader(
        "",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True
    )

    uploaded_images = {}
    image_filenames = []
    if uploaded_images_files:
        for img in uploaded_images_files:
            uploaded_images[img.name] = img.read()
            image_filenames.append(img.name)

    if uploaded_pdf and st.button("Let's Peel"):
        with st.spinner("Generating slides..."):
            pdf_bytes = uploaded_pdf.read()
            result = call_claude_for_slides(
                pdf_bytes,
                st.session_state.layout_info,
                simplification_level,
                image_filenames
            )
            if result:
                st.session_state.slides_json = result
                st.success("PowerPoint generated!")
            else:
                st.error("Failed to generate PowerPoint")

    if "slides_json" in st.session_state and st.session_state.slides_json:
        st.header("Edit Generated Slides")

        slides = st.session_state.slides_json.get("slides", [])

        # Slide selection with dropdown
        st.subheader("Select a Slide to Edit")
        # Create options list with slide numbers and titles
        slide_options = []
        for i, slide in enumerate(slides):
            # Get title from placeholders if it exists
            title = None
            for ph_idx, content in slide.get("placeholders", {}).items():
                placeholder_info = None
                # Find the placeholder info
                for layout in st.session_state.layout_info:
                    if layout['layout_name'] == slide.get("layout_name", ""):
                        for ph in layout['placeholders']:
                            if str(ph['placeholder_idx']) == str(ph_idx):
                                placeholder_info = ph
                                break
                if placeholder_info and "title" in placeholder_info['placeholder_name'].lower():
                    if isinstance(content, dict):
                        title = content.get("text", "")
                    else:
                        title = content
                    break
            
            # Format the option string
            option = f"Slide {i+1}"
            if title:
                option += f" - {title}"
            slide_options.append(option)

        selected_slide_label = st.selectbox(
            "Choose a slide:",
            options=slide_options,
            index=0,
            key="selected_slide_dropdown"
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

                # Text + bullets
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
                if isinstance(content, dict):
                    slides[selected_slide_index]['placeholders'][ph_idx]['text'] = edited_text
                    slides[selected_slide_index]['placeholders'][ph_idx]['bullets'] = edited_bullets
                else:
                    slides[selected_slide_index]['placeholders'][ph_idx] = {
                        "text": edited_text,
                        "bullets": edited_bullets
                    }
            else:
                if isinstance(content, dict):
                    slides[selected_slide_index]['placeholders'][ph_idx]['text'] = edited_text
                else:
                    slides[selected_slide_index]['placeholders'][ph_idx] = edited_text

        st.session_state.slides_json['slides'] = slides

        # Combined Generate and Download button
        prs = Presentation(EXETER_TEMPLATE_PATH)
        prs = create_slides_from_json(prs, st.session_state.slides_json, st.session_state.layout_info, uploaded_images)
        ppt_buffer = BytesIO()
        prs.save(ppt_buffer)
        ppt_buffer.seek(0)
        
        st.download_button(
            "DOWNLOAD",
            data=ppt_buffer.getvalue(),
            file_name="my_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )
        
        # Add space at the bottom
        st.markdown("<div style='padding-bottom: 100px'></div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
