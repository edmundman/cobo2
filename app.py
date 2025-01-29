import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import ChartData, CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_DATA_LABEL_POSITION
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from io import BytesIO
import anthropic
import base64
import json
import fitz
from dotenv import load_dotenv
import os
import re
import matplotlib.pyplot as plt
import uuid
# -------------------- BEGIN: SETUP --------------------
load_dotenv()
st.set_page_config(page_title="Let's Create Your Presentation", layout="wide")

api_key = os.getenv("ANTHROPIC_API_KEY")
if not api_key:
    st.error("Anthropic API key not found. Please check your .env file.")
    st.stop()

client = anthropic.Anthropic(
    api_key=api_key,
    default_headers={"anthropic-beta": "pdfs-2024-09-25"}
)

# Model Name
MODEL_NAME = "claude-3-5-sonnet-20241022"  # <-- your custom model name
GRAPH_PROMPT_PATH = "graphprompt2.txt"     # <-- your prompt file
TEMPLATE_PATH = "Template.pptx"            # <-- your PPT template

# -------------------- Color Palette for Charts --------------------
# Define colors as tuples first for easy conversion
COLOR_VALUES = {
    'primary': (42, 169, 224),     # Light blue #2AA9E0
    'secondary': (10, 75, 117),    # Dark blue #0A4B75
    'accent1': (255, 89, 94),      # Coral red
    'accent2': (86, 204, 157),     # Mint green
    'accent3': (255, 202, 58)      # Yellow
}

# Create PPTX colors
COLORS_PPTX = {
    key: RGBColor(*rgb) for key, rgb in COLOR_VALUES.items()
}

# Create matplotlib colors
COLORS_MPL = [
    '#{:02x}{:02x}{:02x}'.format(r, g, b) for r, g, b in COLOR_VALUES.values()
]

# -------------------- HELPER FUNCTIONS --------------------
def find_shape_in_groups(slide, target_name):
    """Recursively search for a shape with the given name, including in groups."""
    def search_shapes(shapes):
        for shape in shapes:
            if shape.name == target_name:
                return shape
            # If it's a group shape, search within it
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                result = search_shapes(shape.shapes)
                if result:
                    return result
        return None
    
    return search_shapes(slide.shapes)

def _apply_common_chart_style(chart, title=None, text_color=RGBColor(0, 0, 0)):
    """Apply consistent styling to a python-pptx chart."""
    # Chart title
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.text = title
        title_fmt = chart.chart_title.text_frame.paragraphs[0].font
        title_fmt.size = Pt(14)
        title_fmt.name = "Calibri"
        title_fmt.color.rgb = text_color

    # Legend
    if chart.has_legend:
        chart.legend.font.color.rgb = text_color
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False

    # Only apply axis styling for charts that support axes (not pie/donut)
    chart_type = chart.chart_type
    if chart_type not in [XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT]:
        # Category Axis (X-axis)
        try:
            if chart.category_axis:
                chart.category_axis.tick_labels.font.color.rgb = text_color
                if chart.category_axis.has_title:
                    chart.category_axis.axis_title.text_frame.paragraphs[0].font.color.rgb = text_color
        except (ValueError, AttributeError):
            pass

        # Value Axis (Y-axis)
        try:
            if chart.value_axis:
                chart.value_axis.tick_labels.font.color.rgb = text_color
                if chart.value_axis.has_title:
                    chart.value_axis.axis_title.text_frame.paragraphs[0].font.color.rgb = text_color
        except (ValueError, AttributeError):
            pass

    # Data Labels
    try:
        if hasattr(chart, 'plots') and chart.plots:
            for plot in chart.plots:
                if plot.has_data_labels:
                    data_labels = plot.data_labels
                    data_labels.font.color.rgb = text_color
        else:
            for series in chart.series:
                if series.has_data_labels:
                    data_labels = series.data_labels
                    data_labels.font.color.rgb = text_color
    except AttributeError:
        pass
@st.cache_resource
def load_prompt_text(prompt_path):
    with open(prompt_path, "r") as file:
        return file.read()

def display_logo():
    logo_svg = '''
    <div style="text-align: center; margin-top: 2rem;">
        <svg viewBox="0 0 200 80" style="width: 150px;">
            <text x="10" y="60" font-family="Arial" font-size="60" font-weight="bold" fill="#FFFFFF">
                Peel
            </text>
        </svg>
    </div>
    '''
    st.markdown(logo_svg, unsafe_allow_html=True)

def extract_text_from_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    return text

def generate_json_using_claude(prompt, pdf_bytes, simplification_level, progress_callback=None):
    """Sends the prompt and PDF data to Claude to generate JSON data."""
    formatted_prompt = prompt.replace("{{SIMPLIFICATION_LEVEL}}", str(simplification_level))
    pdf_data = base64.b64encode(pdf_bytes).decode('utf-8')

    # Prefix the prompt with "\n\nHuman:" to comply with Claude's requirements
    prefixed_prompt = "\n\nHuman:" + formatted_prompt

    messages = [
        {
            "role": 'user',
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": pdf_data
                    }
                },
                {
                    "type": "text",
                    "text": prefixed_prompt
                }
            ]
        }
    ]

    try:
        if progress_callback:
            progress_callback("Information peeling...", 10)

        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=8192,
            messages=messages,
        )

        if progress_callback:
            progress_callback("Information peeling...", 70)

        assistant_reply = response.content[0].text
        print(assistant_reply)
        json_output_pattern = r"<json_output>\s*(\{.*?\})\s*</json_output>"
        match = re.search(json_output_pattern, assistant_reply, re.DOTALL)

        if match:
            json_string = match.group(1)
            json_data = json.loads(json_string)
        else:
            try:
                json_data = json.loads(assistant_reply)
            except json.JSONDecodeError:
                st.error("Failed to find or parse JSON output in the response.")
                return {}

        charts = json_data.get("CHARTS", [])
        json_data["CHARTS"] = charts

        return json_data

    except Exception as e:
        st.error(f"Error communicating with Claude: {e}")
        return {}
    finally:
        if progress_callback:
            progress_callback("Completed.", 100)

@st.cache_resource
def load_ppt_template(template_path):
    return Presentation(template_path)

def _add_donut_chart(slide, left, top, width, height, chart_title, data, text_color=RGBColor(0, 0, 0)):
    """Add donut chart at specified position using python-pptx."""
    # Check if data is a list of dicts with 'category' and 'value'
    if isinstance(data, list) and all(isinstance(item, dict) for item in data):
        categories = [item.get("category", "") for item in data]
        values = [item.get("value", 0) for item in data]
    else:
        print("Invalid data format for donut chart. Expected a list of dictionaries with 'category' and 'value' keys.")
        categories = ["Default Category"]
        values = [100]
    
    chart_data = ChartData()
    chart_data.categories = categories
    chart_data.add_series("Series 1", values)
    
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT,
        left, top, width, height,
        chart_data
    ).chart
    
    # Make donut hole bigger
    chart.plots[0].donut_hole_size = 60
    
    # Color slices
    series = chart.series[0]
    for idx, point in enumerate(series.points):
        point.format.fill.solid()
        color_key = list(COLORS_PPTX.keys())[idx % len(COLORS_PPTX)]
        point.format.fill.fore_color.rgb = COLORS_PPTX[color_key]
    
    # Data labels
    chart.plots[0].has_data_labels = True
    data_labels = chart.plots[0].data_labels
    data_labels.font.size = Pt(12)
    data_labels.font.bold = True
    data_labels.number_format = '0"%"'
    data_labels.position = XL_DATA_LABEL_POSITION.OUTSIDE_END
    data_labels.font.color.rgb = text_color
    
    _apply_common_chart_style(chart, chart_title, text_color)

def _add_comparison_bars_chart(slide, left, top, width, height, chart_title, data, horizontal=False, text_color=RGBColor(0, 0, 0)):
    """Add a comparison bar chart (horizontal or vertical)."""
    values = data.get("values", [])
    labels = data.get("labels", [])
    if len(values) != len(labels):
        print("Bar chart data mismatch: values and labels must be the same length.")
        return

    pairs = list(zip(labels, values))

    chart_data = CategoryChartData()
    chart_data.categories = [p[0] for p in pairs]
    chart_data.add_series("Series 1", [p[1] for p in pairs])

    chart_type = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    chart = slide.shapes.add_chart(
        chart_type,
        left, top, width, height,
        chart_data
    ).chart

    # Format series
    series = chart.series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = COLORS_PPTX["primary"]

    # Data labels
    series.has_data_labels = True
    data_labels = series.data_labels
    data_labels.font.size = Pt(12)
    data_labels.font.color.rgb = text_color
    data_labels.position = XL_DATA_LABEL_POSITION.OUTSIDE_END

    _apply_common_chart_style(chart, chart_title, text_color)

def _add_patient_count_display(slide, left, top, width, height, chart_title, data, text_color=RGBColor(0, 0, 0)):
    """Display patient count as a prominent text box."""
    count = data.get("count", 0)

    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    # Chart Title
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = chart_title + "\n"
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = text_color

    # Count
    p = tf.add_paragraph()
    run = p.add_run()
    run.text = str(count)
    run.font.size = Pt(36)
    run.font.bold = True
    run.font.color.rgb = COLORS_PPTX["primary"]

def _add_line_chart(slide, left, top, width, height, chart_title, data, text_color=RGBColor(0, 0, 0)):
    """Add a line chart."""
    values = data.get("values", [])
    dates = data.get("dates", [])
    if len(values) != len(dates):
        print("Line chart data mismatch: 'values' and 'dates' must be the same length.")
        return

    pairs = list(zip(dates, values))

    chart_data = CategoryChartData()
    chart_data.categories = [p[0] for p in pairs]
    chart_data.add_series("Series 1", [p[1] for p in pairs])

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        left, top, width, height,
        chart_data
    ).chart

    # Format series
    series = chart.series[0]
    series.format.line.color.rgb = COLORS_PPTX["primary"]
    series.marker.format.fill.solid()
    series.marker.format.fill.fore_color.rgb = COLORS_PPTX["primary"]
    series.marker.size = 7

    # Data labels
    series.has_data_labels = True
    data_labels = series.data_labels
    data_labels.font.size = Pt(12)
    data_labels.font.color.rgb = text_color
    data_labels.position = XL_DATA_LABEL_POSITION.ABOVE

    _apply_common_chart_style(chart, chart_title, text_color)

def _add_stacked_percentage_chart(slide, left, top, width, height, chart_title, data, text_color=RGBColor(0, 0, 0)):
    """Add a stacked percentage chart."""
    categories = data.get("categories", [])
    percentages = data.get("percentages", [])

    if len(categories) != len(percentages):
        print("Stacked percentage chart data mismatch: 'categories' and 'percentages' must be the same length.")
        categories = ["Default Category"]
        percentages = [100]

    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("Series 1", percentages)

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE,
        left, top, width, height,
        chart_data
    ).chart

    # Apply formatting
    chart.has_title = True
    chart.chart_title.text_frame.text = chart_title
    chart.chart_title.text_frame.paragraphs[0].font.size = Pt(14)
    chart.chart_title.text_frame.paragraphs[0].font.name = "Calibri"
    chart.chart_title.text_frame.paragraphs[0].font.color.rgb = text_color

    # Color slices
    series = chart.series[0]
    for idx, point in enumerate(series.points):
        point.format.fill.solid()
        color_key = list(COLORS_PPTX.keys())[idx % len(COLORS_PPTX)]
        point.format.fill.fore_color.rgb = COLORS_PPTX[color_key]

    # Data labels
    chart.plots[0].has_data_labels = True
    data_labels = chart.plots[0].data_labels
    data_labels.font.size = Pt(12)
    data_labels.font.bold = True
    data_labels.number_format = '0"%"'
    data_labels.position = XL_DATA_LABEL_POSITION.OUTSIDE_END
    data_labels.font.color.rgb = text_color

    _apply_common_chart_style(chart, chart_title, text_color)

def populate_ppt_template(json_data, prs, uploaded_images):
    """
    Insert text into existing slides placeholders AND replace graph placeholders with python-pptx charts.
    Additionally, replace image placeholders with user-uploaded images if provided.
    """
    # 1) Insert text into placeholders
    for slide_number, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                text_frame = shape.text_frame
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.text = run.text.replace("(Title)", json_data.get("Title", ""))
                        run.text = run.text.replace("(AUTHOR_NAMES)", ", ".join(json_data.get("AUTHOR_NAMES", [])))
                        run.text = run.text.replace("(PAPER_PMID)", json_data.get("PAPER_PMID", ""))
                        run.text = run.text.replace("(PAPER_DOI)", json_data.get("PAPER_DOI", ""))
                        run.text = run.text.replace("(background)", json_data.get("Background_Info", ""))
                        run.text = run.text.replace("(Patient Quote)", json_data.get("Patient_Quote", ""))
                        run.text = run.text.replace("(Patient name)", json_data.get("Patient_Name", ""))
                        run.text = run.text.replace("(Date)", json_data.get("Date", ""))
                        run.text = run.text.replace("(AIMS)", "\n".join(json_data.get("AIMS", [])))
                        run.text = run.text.replace("(Methods)", json_data.get("Methods", ""))
                        run.text = run.text.replace("(Findings)", "\n".join(json_data.get("Findings", [])))
                        run.text = run.text.replace("(Conclusion)", json_data.get("Conclusion", ""))
                        run.text = run.text.replace("(Slide_Number)", str(slide_number))

    # 2) Replace graph placeholders with charts
    charts_list = json_data.get("CHARTS", [])
    if not charts_list:
        print("No charts found in JSON data.")
    else:
        # Extract all 'graph_' placeholders across all slides and group them by name
        placeholder_dict = {}
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.name.startswith("graph_"):
                    if shape.name not in placeholder_dict:
                        placeholder_dict[shape.name] = []
                    placeholder_dict[shape.name].append(shape)

        # Determine the total number of placeholders and sort them
        total_placeholders = sorted(placeholder_dict.keys(), key=lambda x: int(x.split('_')[1]))

        # Iterate over placeholders and assign charts
        for idx, placeholder_name in enumerate(total_placeholders, start=1):
            # Determine which chart to use
            if idx <= len(charts_list):
                chart_info = charts_list[idx - 1]
            elif len(charts_list) >= 2:
                chart_info = charts_list[1]  # Use Chart 2 as replacement
            elif len(charts_list) >= 1:
                chart_info = charts_list[0]  # Use Chart 1 as fallback
            else:
                print(f"No charts available to replace placeholder '{placeholder_name}'.")
                continue

            # Get all shapes with this placeholder name
            shapes = placeholder_dict[placeholder_name]
            for shape in shapes:
                slide = shape.part.slide
                left = shape.left
                top = shape.top
                width = shape.width
                height = shape.height
                chart_title = chart_info.get("chart_title", "")
                data = chart_info.get("data", {})

                # Determine the slide number for text color
                slide_number = prs.slides.index(slide) + 1
                text_color = RGBColor(255, 255, 255) if slide_number == 7 else RGBColor(0, 0, 0)

                # Remove the old placeholder shape
                sp = shape._element
                sp.getparent().remove(sp)

                # Add the new chart based on chart_type
                chart_type = chart_info.get("chart_type", "")
                if chart_type == "donut":
                    _add_donut_chart(slide, left, top, width, height, chart_title, data, text_color)
                elif chart_type == "comparison_bars":
                    _add_comparison_bars_chart(slide, left, top, width, height, chart_title, data, horizontal=False, text_color=text_color)
                elif chart_type == "trend_line":
                    _add_line_chart(slide, left, top, width, height, chart_title, data, text_color)
                elif chart_type == "stacked_percentage":
                    _add_stacked_percentage_chart(slide, left, top, width, height, chart_title, data, text_color)
                elif chart_type == "patient_count":
                    _add_patient_count_display(slide, left, top, width, height, chart_title, data, text_color)
                else:
                    print(f"Chart type '{chart_type}' is not currently supported.")

    # 3) Replace image placeholders with user-uploaded images
    image_mappings = {
        "image1": {"page": 1, "name": "Image 1"},
        "image2": {"page": 2, "name": "Image 2"},
        "image3": {"page": 5, "name": "Image 3"},
        "image4": {"page": 6, "name": "Image 4"},
        "image5": {"page": 8, "name": "Image 5"},
    }

    for image_key, info in image_mappings.items():
        uploaded_image = uploaded_images.get(image_key)
        if uploaded_image:
            found = False
            for slide in prs.slides:
                if prs.slides.index(slide) + 1 == info["page"]:
                    # Use the new function to find the shape, including in groups
                    shape = find_shape_in_groups(slide, image_key)
                    if shape:
                        found = True
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height
                        
                        # If shape is in a group, store the group and shape's index
                        parent_group = shape.parent_group if hasattr(shape, 'parent_group') else None
                        shape_index = shape.element.getnext().index if parent_group else None
                        
                        # Remove the old image
                        sp = shape._element
                        sp.getparent().remove(sp)
                        
                        # Add the new image
                        image_stream = BytesIO(uploaded_image)
                        new_picture = slide.shapes.add_picture(image_stream, left, top, width, height)
                        
                        # If the original was in a group, move the new picture to the group
                        if parent_group:
                            new_picture._element.addnext(parent_group._element[shape_index])
                        
                        print(f"Replaced {info['name']} on Page {info['page']} with user-uploaded image.")
                        st.write(f"Replaced {info['name']} on Page {info['page']} with user-uploaded image.")
                        break
            if not found:
                print(f"Placeholder '{image_key}' not found on Page {info['page']}.")
                print(f"Placeholder '{image_key}' not found on Page {info['page']}.")

    # 4) Return the final PPT as a stream
    ppt_stream = BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

def show_chart_previews(json_data):
    """Display previews of the charts defined in the JSON data."""
    st.header("Chart Previews")
    for idx, chart in enumerate(json_data.get("CHARTS", []), start=1):
        st.subheader(f"Chart #{idx}: {chart.get('chart_title', 'No Title')}")
        chart_type = chart.get("chart_type", "")
        data = chart.get("data", {})
        if chart_type == "donut":
            labels = [item["category"] for item in data]
            sizes = [item["value"] for item in data]
            fig, ax = plt.subplots()
            ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140, colors=COLORS_MPL[:len(labels)])
            ax.axis('equal')
            st.pyplot(fig)

        elif chart_type == "comparison_bars":
            labels = data.get("labels", [])
            values = data.get("values", [])
            fig, ax = plt.subplots()
            ax.bar(labels, values, color=COLORS_MPL[0])
            ax.set_title(chart.get("chart_title", "Bar Chart"))
            ax.set_ylabel("Values")
            st.pyplot(fig)

        elif chart_type == "trend_line":
            categories = data.get("dates", [])
            values = data.get("values", [])
            fig, ax = plt.subplots()
            ax.plot(categories, values, marker='o', color=COLORS_MPL[0])
            ax.set_title(chart.get("chart_title", "Line Chart"))
            ax.set_xlabel("Dates")
            ax.set_ylabel("Values")
            st.pyplot(fig)

        elif chart_type == "stacked_percentage":
            categories = data.get("categories", [])
            percentages = data.get("percentages", [])
            fig, ax = plt.subplots()
            ax.pie(percentages, labels=categories, autopct='%1.1f%%', startangle=140, colors=COLORS_MPL[:len(categories)])
            ax.axis('equal')
            ax.set_title(chart.get("chart_title", "Stacked Percentage Chart"))
            st.pyplot(fig)

        elif chart_type == "patient_count":
            count = data.get("count", 0)
            fig, ax = plt.subplots(figsize=(4, 3))
            ax.text(0.5, 0.5, str(count), fontsize=48, ha='center', va='center', color=COLORS_MPL[0])
            ax.set_title(chart.get("chart_title", "Patient Count"), fontsize=14)
            ax.axis('off')
            st.pyplot(fig)

def main():
    display_logo()
    st.markdown('<h1 style="text-align: center; color: white;">Let’s Create Your Presentation</h1>', unsafe_allow_html=True)


    # Session states
    if "json_data" not in st.session_state:
        st.session_state.json_data = None
    if "ppt_file" not in st.session_state:
        st.session_state.ppt_file = None
    if "current_section" not in st.session_state:
        st.session_state.current_section = "Page 1: Title and Author Names"
    if "uploaded_images" not in st.session_state:
        st.session_state.uploaded_images = {
            "image1": None,
            "image2": None,
            "image3": None,
            "image4": None,
            "image5": None
        }

    prompt_text = load_prompt_text(GRAPH_PROMPT_PATH)
    uploaded_pdf = st.file_uploader("Upload your PDF file", type="pdf")

    # Simplification level
    simplification_level = st.select_slider(
        "Select Simplification Level",
        options=list(range(1, 11)),
        format_func=lambda x: "Academic" if x == 1 else "Patient" if x == 10 else f"Level {x}",
        value=5
    )

    if uploaded_pdf:
        pdf_bytes = uploaded_pdf.read()

        if st.button("Let's Peel"):
            progress_bar = st.progress(0)
            status_text = st.empty()

            def update_progress(step_description, pct):
                progress_bar.progress(pct)
                status_text.text(step_description)

            st.session_state.json_data = generate_json_using_claude(
                prompt_text,
                pdf_bytes,
                simplification_level,
                progress_callback=update_progress
            )
            if st.session_state.json_data:
                progress_bar.progress(100)
                status_text.text("Peeling Complete!")
                st.success("Initial Peeling Complete!")

    # If we have JSON in the session
    if st.session_state.json_data:
        sections = [
            "Page 1: Title and Author Names",
            "Page 2: Background",
            "Page 3: Patient Quote",
            "Page 4: Key Statistics",
            "Page 5: Aims",
            "Page 6: Methods",
            "Page 7: Results",
            "Page 8: Conclusions"
        ]
        st.session_state.current_section = st.radio(
            "Select section to edit:",
            sections,
            key="section_selector"
        )

        # Let user edit
        json_data_before = st.session_state.json_data.copy()
        st.session_state.json_data = edit_json_section(
            st.session_state.json_data,
            st.session_state.current_section
        )

        # Let user preview charts
        if st.checkbox("Show Chart Previews"):
            try:
                show_chart_previews(st.session_state.json_data)
            except AttributeError as e:
                st.error(f"Error in chart previews: {e}")

        # Generate PPT
        if st.button("Generate PowerPoint"):
            # Create progress indicators
            ppt_progress_bar = st.progress(0)
            ppt_status_text = st.empty()
            ppt_status_text.text("Generating PowerPoint...")

            try:
                # Generate a fresh Presentation each time
                fresh_prs = load_ppt_template(TEMPLATE_PATH)
                new_ppt_stream = populate_ppt_template(
                    st.session_state.json_data,
                    fresh_prs,
                    st.session_state.uploaded_images
                )
                # Store in session_state
                st.session_state.ppt_file = new_ppt_stream
                
                ppt_progress_bar.progress(100)
                ppt_status_text.text("PowerPoint Generated!")
                st.success("PowerPoint generated successfully!")
            except Exception as e:
                st.error(f"Error generating PowerPoint: {str(e)}")
                ppt_status_text.text("Error generating PowerPoint")
                return

        # If PPT is generated, show download button
        if st.session_state.ppt_file:
            st.download_button(
                label="Download PowerPoint",
                data=st.session_state.ppt_file,
                file_name=f"presentation_{uuid.uuid4()}.pptx",  # Unique filename
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key=f"download_ppt_{uuid.uuid4()}"             # Optional unique key
            )

def edit_json_section(json_data, section_name):
    """Let user edit each JSON section in Streamlit."""
    image_mappings = {
        "Page 1: Title and Author Names": {"key": "image1", "label": "Image 1", "page": 1},
        "Page 2: Background": {"key": "image2", "label": "Image 2", "page": 2},
        "Page 5: Aims": {"key": "image3", "label": "Image 3", "page": 5},
        "Page 6: Methods": {"key": "image4", "label": "Image 4", "page": 6},
        "Page 8: Conclusions": {"key": "image5", "label": "Image 5", "page": 8},
    }

    if section_name == "Page 1: Title and Author Names":
        st.header("Page 1: Title and Author Names")
        st.text_input("Title", value=json_data.get("Title", ""), key="title")
        authors_str = ", ".join(json_data.get("AUTHOR_NAMES", []))
        authors_str = st.text_input("Authors (comma-separated)", value=authors_str, key="authors")
        st.text_input("DOI", value=json_data.get("PAPER_DOI", ""), key="doi")

        json_data["Title"] = st.session_state.title
        json_data["AUTHOR_NAMES"] = [a.strip() for a in st.session_state.authors.split(",") if a.strip()]
        json_data["PAPER_DOI"] = st.session_state.doi

        # Image Uploader for Page 1
        image_info = image_mappings.get(section_name)
        if image_info:
            uploaded_image = st.file_uploader(
                f"Upload {image_info['label']} (Page {image_info['page']})",
                type=["png", "jpg", "jpeg"],
                key=image_info["key"]
            )
            if uploaded_image:
                st.session_state.uploaded_images[image_info["key"]] = uploaded_image.read()
                st.success(f"{image_info['label']} uploaded successfully.")
            else:
                if st.session_state.uploaded_images.get(image_info["key"]):
                    st.image(st.session_state.uploaded_images[image_info["key"]], caption=image_info['label'])

    elif section_name == "Page 2: Background":
        st.header("Page 2: Background")
        st.text_area("Background Info", value=json_data.get("Background_Info", ""), key="background")
        json_data["Background_Info"] = st.session_state.background

        # Image Uploader for Page 2
        image_info = image_mappings.get(section_name)
        if image_info:
            uploaded_image = st.file_uploader(
                f"Upload {image_info['label']} (Page {image_info['page']})",
                type=["png", "jpg", "jpeg"],
                key=image_info["key"]
            )
            if uploaded_image:
                st.session_state.uploaded_images[image_info["key"]] = uploaded_image.read()
                st.success(f"{image_info['label']} uploaded successfully.")
            else:
                if st.session_state.uploaded_images.get(image_info["key"]):
                    st.image(st.session_state.uploaded_images[image_info["key"]], caption=image_info['label'])

    elif section_name == "Page 3: Patient Quote":
        st.header("Page 3: Patient Quote")
        st.text_area("Patient Quote", value=json_data.get("Patient_Quote", ""), key="patient_quote")
        st.text_input("Patient Name", value=json_data.get("Patient_Name", ""), key="patient_name")
        json_data["Patient_Quote"] = st.session_state.patient_quote
        json_data["Patient_Name"] = st.session_state.patient_name

    elif section_name == "Page 4: Key Statistics":
        st.header("Page 4: Key Statistics")
        charts = json_data.get("CHARTS", [])
        for idx, chart in enumerate(charts, start=1):
            chart_title = chart.get("chart_title", f"Chart {idx}")
            chart_type = chart.get("chart_type", "unknown")
            st.subheader(f"Graph {idx}: {chart_title} ({chart_type})")

            if chart_type == "donut":
                st.markdown(f"**{chart_title}**")
                categories = []
                values = []
                for item in chart.get("data", []):
                    category = st.text_input(f"Category {item.get('category', '')}", value=item.get("category", ""), key=f"graph_{idx}_category_{item.get('category', '')}")
                    value = st.number_input(
                        f"Value for {category}",
                        value=float(item.get("value", 0.0)),
                        min_value=0.0,
                        max_value=100.0,
                        step=0.1,
                        key=f"graph_{idx}_value_{item.get('category', '')}"
                    )
                    categories.append(category)
                    values.append(value)
                json_data["CHARTS"][idx - 1]["data"] = [{"category": c, "value": v} for c, v in zip(categories, values)]

            elif chart_type == "comparison_bars":
                st.markdown(f"**{chart_title}**")
                labels = []
                values = []
                for label, value in zip(chart.get("data", {}).get("labels", []), chart.get("data", {}).get("values", [])):
                    new_label = st.text_input(f"Label for {label}", value=label, key=f"graph_{idx}_label_{label}")
                    new_value = st.number_input(
                        f"Value for {new_label}",
                        value=float(value),
                        min_value=0.0,
                        max_value=100.0,
                        step=0.1,
                        key=f"graph_{idx}_value_{label}"
                    )
                    labels.append(new_label)
                    values.append(new_value)
                json_data["CHARTS"][idx - 1]["data"] = {
                    "labels": labels,
                    "values": values
                }

            elif chart_type == "patient_count":
                st.markdown(f"**{chart_title}**")
                count = st.number_input(
                    f"Count for {chart_title}",
                    value=int(chart.get("data", {}).get("count", 0)),
                    min_value=0,
                    step=1,
                    key=f"graph_{idx}_count"
                )
                json_data["CHARTS"][idx - 1]["data"] = {
                    "count": count
                }

            elif chart_type == "trend_line":
                st.markdown(f"**{chart_title}**")
                dates = []
                values = []
                for date, value in zip(chart.get("data", {}).get("dates", []), chart.get("data", {}).get("values", [])):
                    new_date = st.text_input(f"Date for {date}", value=date, key=f"graph_{idx}_date_{date}")
                    new_value = st.number_input(
                        f"Value for {new_date}",
                        value=float(value),
                        min_value=0.0,
                        max_value=100.0,
                        step=0.1,
                        key=f"graph_{idx}_value_{date}"
                    )
                    dates.append(new_date)
                    values.append(new_value)
                json_data["CHARTS"][idx - 1]["data"] = {
                    "dates": dates,
                    "values": values
                }

            elif chart_type == "stacked_percentage":
                st.markdown(f"**{chart_title}**")
                categories = []
                percentages = []
                for i, (category, percentage) in enumerate(zip(chart.get("data", {}).get("categories", []), chart.get("data", {}).get("percentages", [])), start=1):
                    new_category = st.text_input(f"Category {i}", value=category, key=f"graph_{idx}_category_{i}")
                    new_percentage = st.number_input(
                        f"Percentage for {new_category}",
                        value=float(percentage),
                        min_value=0.0,
                        max_value=100.0,
                        step=0.1,
                        key=f"graph_{idx}_percentage_{i}"
                    )
                    categories.append(new_category)
                    percentages.append(new_percentage)
                json_data["CHARTS"][idx - 1]["data"] = {
                    "categories": categories,
                    "percentages": percentages
                }

            else:
                print(f"Chart type '{chart_type}' is not supported for editing.")

    elif section_name == "Page 5: Aims":
        st.header("Page 5: Aims")
        st.text_area("Aims (one per line)", value="\n".join(json_data.get("AIMS", [])), key="aims")
        json_data["AIMS"] = [a.strip() for a in st.session_state.aims.split("\n") if a.strip()]

        # Image Uploader for Page 5
        image_info = image_mappings.get(section_name)
        if image_info:
            uploaded_image = st.file_uploader(
                f"Upload {image_info['label']} (Page {image_info['page']})",
                type=["png", "jpg", "jpeg"],
                key=image_info["key"]
            )
            if uploaded_image:
                st.session_state.uploaded_images[image_info["key"]] = uploaded_image.read()
                st.success(f"{image_info['label']} uploaded successfully.")
            else:
                if st.session_state.uploaded_images.get(image_info["key"]):
                    st.image(st.session_state.uploaded_images[image_info["key"]], caption=image_info['label'])

    elif section_name == "Page 6: Methods":
        st.header("Page 6: Methods")
        st.text_area("Methods", value=json_data.get("Methods", ""), key="methods")
        json_data["Methods"] = st.session_state.methods

        # Image Uploader for Page 6
        image_info = image_mappings.get(section_name)
        if image_info:
            uploaded_image = st.file_uploader(
                f"Upload {image_info['label']} (Page {image_info['page']})",
                type=["png", "jpg", "jpeg"],
                key=image_info["key"]
            )
            if uploaded_image:
                st.session_state.uploaded_images[image_info["key"]] = uploaded_image.read()
                st.success(f"{image_info['label']} uploaded successfully.")
            else:
                if st.session_state.uploaded_images.get(image_info["key"]):
                    st.image(st.session_state.uploaded_images[image_info["key"]], caption=image_info['label'])

    elif section_name == "Page 7: Results":
        st.header("Page 7: Results")
        st.text_area("Findings (one per line)", value="\n".join(json_data.get("Findings", [])), key="findings")
        json_data["Findings"] = [f.strip() for f in st.session_state.findings.split("\n") if f.strip()]

    elif section_name == "Page 8: Conclusions":
        st.header("Page 8: Conclusions")
        st.text_area("Conclusion", value=json_data.get("Conclusion", ""), key="conclusion")
        json_data["Conclusion"] = st.session_state.conclusion

        # Image Uploader for Page 8
        image_info = image_mappings.get(section_name)
        if image_info:
            uploaded_image = st.file_uploader(
                f"Upload {image_info['label']} (Page {image_info['page']})",
                type=["png", "jpg", "jpeg"],
                key=image_info["key"]
            )
            if uploaded_image:
                st.session_state.uploaded_images[image_info["key"]] = uploaded_image.read()
                st.success(f"{image_info['label']} uploaded successfully.")
            else:
                if st.session_state.uploaded_images.get(image_info["key"]):
                    st.image(st.session_state.uploaded_images[image_info["key"]], caption=image_info['label'])

    return json_data

if __name__ == "__main__":
    main()
