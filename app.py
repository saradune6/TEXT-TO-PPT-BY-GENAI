import streamlit as st
import base64
import google.generativeai as genai
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import os
from dotenv import load_dotenv

load_dotenv()

genai.configure(api_key=os.getenv('GOOGLE_API_KEY'))  # Replace with your actual Google API key

# Define custom formatting options
TITLE_FONT_SIZE = Pt(30)
SLIDE_FONT_SIZE = Pt(16)
MAX_CONTENT_LENGTH = 200  # Define the maximum number of characters per slide

# Define margin size for the content box (in inches)
MARGIN_LEFT = Inches(1)
MARGIN_TOP = Inches(1.5)  # Increased top margin to avoid overlap with the title
MARGIN_RIGHT = Inches(1)
MARGIN_BOTTOM = Inches(1)

def generate_slide_titles(topic):
    prompt = f"Generate 5 slide titles for the topic '{topic}'."
    model = genai.GenerativeModel('gemini-pro')
    response = model.generate_content(prompt)
    return response.text.split("\n")

def generate_slide_content(slide_title):
    prompt = f"Generate content for the slide: '{slide_title}'."
    model = genai.GenerativeModel('gemini-pro')
    response = model.generate_content(prompt)
    return response.text

def apply_theme(slide, theme):
    """
    Apply a selected theme to the slide, setting colors and fonts.
    """
    if theme == "Light":
        background_color = RGBColor(242, 242, 242)  # Light gray background
        title_color = RGBColor(0, 0, 0)  # Black title
        content_color = RGBColor(50, 50, 50)  # Dark gray content text
    elif theme == "Dark":
        background_color = RGBColor(0, 0, 0)  # Black background
        title_color = RGBColor(255, 255, 255)  # White title
        content_color = RGBColor(200, 200, 200)  # Light gray content text
    elif theme == "Blue":
        background_color = RGBColor(0, 102, 204)  # Blue background
        title_color = RGBColor(255, 255, 255)  # White title
        content_color = RGBColor(255, 255, 255)  # White content text
    else:
        background_color = RGBColor(255, 255, 255)  # Default white background
        title_color = RGBColor(0, 0, 0)  # Black title
        content_color = RGBColor(50, 50, 50)  # Dark gray content text

    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = background_color

    for shape in slide.shapes:
        if shape.has_text_frame:
            text_frame = shape.text_frame
            for paragraph in text_frame.paragraphs:
                if shape == slide.shapes.title:
                    paragraph.font.size = TITLE_FONT_SIZE
                    paragraph.font.color.rgb = title_color
                else:
                    paragraph.font.size = SLIDE_FONT_SIZE
                    paragraph.font.color.rgb = content_color

def create_presentation(topic, slide_titles, slide_contents, theme):
    # Check if the directory exists, if not, create it
    if not os.path.exists('generated_ppt'):
        os.makedirs('generated_ppt')

    prs = pptx.Presentation()
    slide_layout = prs.slide_layouts[1]  # 1 is for title + content slide

    # Add title slide with custom theme
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = topic
    apply_theme(title_slide, theme)

    for slide_title, slide_content in zip(slide_titles, slide_contents):
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = slide_title
        apply_theme(slide, theme)

        # Limit content length and adjust font size
        if len(slide_content) > MAX_CONTENT_LENGTH:
            slide_content = slide_content[:MAX_CONTENT_LENGTH] + "..."
        
        # Add content to the slide
        content_box = slide.shapes.placeholders[1]
        content_box.text = slide_content

        # Center the content with margins, ensuring no overlap with title
        content_box.left = MARGIN_LEFT  # Left margin
        content_box.top = MARGIN_TOP   # Top margin (adjusted to avoid title overlap)
        content_box.width = Inches(10) - MARGIN_LEFT - MARGIN_RIGHT  # Full width minus left and right margins
        content_box.height = Inches(7.5) - MARGIN_TOP - MARGIN_BOTTOM  # Full height minus top and bottom margins

    prs.save(f"generated_ppt/{topic}_presentation.pptx")


def main():
    st.title("PowerPoint Presentation Generator with Google Gemini")

    topic = st.text_input("Enter the topic for your presentation:")
    theme = st.selectbox("Select a theme for your presentation:", ["Light", "Dark", "Blue", "Default"])
    generate_button = st.button("Generate Presentation")

    if generate_button and topic:
        st.info("Generating presentation... Please wait.")
        slide_titles = generate_slide_titles(topic)
        filtered_slide_titles = [item for item in slide_titles if item.strip() != '']
        slide_contents = [generate_slide_content(title) for title in filtered_slide_titles]
        create_presentation(topic, filtered_slide_titles, slide_contents, theme)
        
        st.success("Presentation generated successfully!")
        st.markdown(get_ppt_download_link(topic), unsafe_allow_html=True)

def get_ppt_download_link(topic):
    ppt_filename = f"generated_ppt/{topic}_presentation.pptx"

    with open(ppt_filename, "rb") as file:
        ppt_contents = file.read()

    b64_ppt = base64.b64encode(ppt_contents).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64_ppt}" download="{ppt_filename}">Download the PowerPoint Presentation</a>'


if __name__ == "__main__":
    main()
