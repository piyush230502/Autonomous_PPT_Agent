import streamlit as st
import os
from dotenv import load_dotenv
from crewai import Agent, Task, Crew, Process
from crewai_tools import SerperDevTool
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from io import BytesIO
import re # For parsing the output

# --- Configuration & API Key Handling ---

# Try loading .env file for local development
try:
    load_dotenv()
except Exception as e:
    # No warning needed here, get_api_key will handle missing keys
    pass

# Function to get API key securely
def get_api_key(service_name, key_name):
    """Gets API key from Streamlit secrets or environment variables."""
    # First, try Streamlit secrets (preferred for deployment)
    key = st.secrets.get(key_name)
    if key:
        # st.sidebar.success(f"{service_name} key loaded from secrets.", icon="✅") # Optional: Confirmation
        return key
    # Fallback to environment variable (useful for local dev with .env)
    key = os.environ.get(key_name)
    if key:
        # st.sidebar.success(f"{service_name} key loaded from environment.", icon="✅") # Optional: Confirmation
        return key
    # Key not found - return None. The UI will handle showing a warning.
    return None

# --- LLM Configuration ---

def get_llm(provider, api_key):
    """Initializes and returns the selected LLM."""
    # The check for api_key is now done before calling this function
    # if not api_key:
    #     st.error(f"API Key for {provider} is missing. Please add it to your secrets or enter it.")
    #     return None
    try:
        if provider == "Groq":
            from langchain_groq import ChatGroq
            # Ensure GROQ_API_KEY is set as an environment variable
            # for underlying libraries like LiteLLM to potentially pick up.
            os.environ["GROQ_API_KEY"] = api_key # Setting env var here might be redundant if get_api_key already found it in env
            # Pass the key explicitly to ChatGroq as well
            
            return ChatGroq(api_key=api_key, model_name="groq/llama3-8b-8192", temperature=0.7) # Or llama3-70b-8192

        elif provider == "Google Gemini":
            from langchain_google_genai import ChatGoogleGenerativeAI
            # Ensure GOOGLE_API_KEY is set as env var for the library
            os.environ["GOOGLE_API_KEY"] = api_key
            return ChatGoogleGenerativeAI(model="gemini-1.5-flash", api_key=api_key, temperature=0.7) # Or gemini-pro

        elif provider == "OpenAI":
            from langchain_openai import ChatOpenAI
            # Similarly, ensure OPENAI_API_KEY is set if needed by underlying libs,
            # although ChatOpenAI usually handles it well with the api_key param.
            os.environ["OPENAI_API_KEY"] = api_key
            return ChatOpenAI(api_key=api_key, model_name="gpt-3.5-turbo", temperature=0.7) # Or gpt-4o
        else:
            st.error(f"Unsupported LLM provider: {provider}")
            return None
    except ImportError as e:
        st.error(f"Failed to import library for {provider}. Please ensure it's installed (`pip install langchain-{provider.lower()}`). Error: {e}")
        return None
    except Exception as e:
        st.error(f"Failed to initialize LLM for {provider}. Error: {e}")
        # Consider adding more detailed logging here if needed
        # import traceback
        # st.error(traceback.format_exc())
        return None
    # Removed the print statement that showed the LLM object
    # print("🔍 ChatGroq LLM initialized:", llm)

# --- PowerPoint Generation ---

def parse_crew_output(text):
    """
    Parses the text output from CrewAI into a structured list of slides.
    Assumes a simple format like:
    SLIDE: Title of Slide 1
    - Bullet point 1
    - Bullet point 2

    SLIDE: Title of Slide 2
    - Point A
    - Point B
    """
    slides = []
    # Split the text into potential slide blocks based on "SLIDE:"
    slide_blocks = re.split(r'\n*SLIDE:\s*', text.strip())

    for block in slide_blocks:
        if not block.strip():
            continue

        lines = block.strip().split('\n')
        title = lines[0].strip()
        content_points = [line.strip('- ').strip() for line in lines[1:] if line.strip() and line.strip().startswith('-')]

        if title: # Only add if a title was found
             slides.append({"title": title, "content": content_points})
        elif content_points: # Handle case where title might be missing but content exists
             slides.append({"title": "Content Slide", "content": content_points}) # Default title

    # Handle case where no "SLIDE:" marker was found, treat the whole output as one slide
    if not slides and text.strip():
        lines = text.strip().split('\n')
        # Try to find a title (first line?) or use a default
        title = lines[0].strip() if lines else "Generated Content"
        content_points = [line.strip('- ').strip() for line in lines[1:] if line.strip()]
        if not content_points and len(lines) > 0: # If no bullet points, maybe it's just text
             content_points = lines[1:] # Take remaining lines as content
        elif not content_points and len(lines) == 1: # Only one line found
             content_points = [lines[0]] # Use the title line as content too
             title = "Generated Content"

        slides.append({"title": title, "content": content_points})


    # Add a default title slide if the first slide has no title or content is minimal
    # if not slides or not slides[0]["title"] or len(slides[0]["content"]) == 0:
    #      slides.insert(0, {"title": "Presentation Title", "content": ["Generated by AI"]})
    # Let's rely on the create_ppt function to add the title slide instead

    return slides


def create_ppt(slides_data, topic):
    """Creates a PowerPoint presentation from structured slide data."""
    prs = Presentation()
    # Use a built-in layout: Title and Content
    title_content_layout = prs.slide_layouts[1]
    # Use title slide layout
    title_slide_layout = prs.slide_layouts[0]

    # Add Title Slide
    slide = prs.slides.add_slide(title_slide_layout)
    title_placeholder = slide.shapes.title
    subtitle_placeholder = slide.placeholders[1] # Placeholder index 1 is typically the subtitle
    title_placeholder.text = topic.title() if topic else "AI Generated Presentation"
    subtitle_placeholder.text = "Created using CrewAI and Streamlit"

    # Add Content Slides
    for i, slide_info in enumerate(slides_data):
        # Skip adding empty slides potentially created by parsing issues
        if not slide_info.get("title") and not slide_info.get("content"):
            continue

        slide = prs.slides.add_slide(title_content_layout)
        title = slide.shapes.title
        body = slide.placeholders[1] # Placeholder index 1 is typically the Content box

        title.text = slide_info.get("title", f"Slide {i+2}") # Use i+2 because title slide is 1

        # Add content points to the body placeholder
        tf = body.text_frame
        tf.clear() # Clear existing text/bullets
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE # Adjust text size to fit

        content_points = slide_info.get("content", [])
        if content_points:
            # Add the first point without indentation
            p = tf.add_paragraph()
            p.text = content_points[0]
            p.font.size = Pt(18) # Adjust font size as needed
            # Add subsequent points with indentation
            for point in content_points[1:]:
                p = tf.add_paragraph()
                p.text = point
                p.level = 1 # Indent level
                p.font.size = Pt(18)
        else:
            p = tf.add_paragraph()
            p.text = "No specific points generated for this slide."
            p.font.size = Pt(18)

    # Save to a BytesIO object
    ppt_io = BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# --- CrewAI Setup ---

def setup_and_run_crew(topic, llm, serper_api_key):
    """Defines agents, tasks, and runs the Crew."""
    # Key check is now done before calling this function
    # if not serper_api_key:
    #     st.error("Serper API Key is missing. Cannot perform web searches.")
    #     return None

    try:
        search_tool = SerperDevTool(api_key=serper_api_key)
    except Exception as e:
        st.error(f"Failed to initialize SerperDevTool: {e}. Check your API key (must be in secrets or .env).")
        return None

    # Define Agents
    researcher = Agent(
        role='Senior Research Analyst',
        goal=f'Gather comprehensive and relevant information on the topic: {topic}',
        backstory=(
            "You are an expert research analyst with a knack for finding "
            "the most current and credible information from the web. "
            "You are skilled at synthesizing information into concise summaries."
        ),
        verbose=True,
        allow_delegation=False,
        tools=[search_tool],
        llm=llm
    )

    outliner = Agent(
        role='Presentation Outline Specialist',
        goal=f'Create a logical and engaging presentation outline for the topic: {topic}',
        backstory=(
            "You are a master at structuring information. You take research findings "
            "and organize them into a clear, compelling presentation flow, "
            "defining slide titles and key areas to cover."
        ),
        verbose=True,
        allow_delegation=False,
        llm=llm
    )

    writer = Agent(
        role='Content Writer for Presentations',
        goal=f'Write concise and informative content (bullet points) for each slide defined in the outline for the topic: {topic}',
        backstory=(
            "You specialize in crafting clear, concise, and engaging content "
            "specifically for PowerPoint presentations. You transform outlines and research "
            "into easy-to-understand bullet points suitable for slides. "
            "You focus on clarity and brevity."
        ),
        verbose=True,
        allow_delegation=False,
        llm=llm
    )

    # Define Tasks
    research_task = Task(
        description=(
            f"Conduct thorough research on '{topic}'. Identify key facts, "
            "statistics, relevant examples, and potential subtopics. "
            "Compile your findings into a detailed report."
        ),
        expected_output="A comprehensive report summarizing the key information found about the topic.",
        agent=researcher
    )

    outline_task = Task(
        description=(
            f"Based on the research report for '{topic}', create a logical presentation outline. "
            "The outline should include a title slide suggestion and a sequence of content slides, "
            "each with a clear title indicating its topic. Aim for 5-10 content slides."
        ),
        expected_output="A structured presentation outline with slide titles.",
        agent=outliner,
        context=[research_task] # Depends on the research task
    )

    content_task = Task(
        description=(
            f"Using the research report and the presentation outline for '{topic}', write the content for each slide. "
            "For each slide title in the outline, generate 3-5 concise bullet points summarizing the key information. "
            "Focus on clarity and make sure the points directly relate to the slide title. "
            "Format the output clearly, starting each slide's content with 'SLIDE: [Slide Title]' followed by bullet points starting with '-'. "
            "Example:\nSLIDE: Introduction to {topic}\n- Point 1\n- Point 2\n\nSLIDE: Key Benefit 1\n- Detail A\n- Detail B"
        ),
        expected_output=(
            "The final presentation content, formatted with 'SLIDE: [Title]' markers and bullet points for each slide."
        ),
        agent=writer,
        context=[research_task, outline_task] # Depends on research and outline
    )

    # Create and Run the Crew
    presentation_crew = Crew(
        agents=[researcher, outliner, writer],
        tasks=[research_task, outline_task, content_task],
        process=Process.sequential,
        verbose=True # Logs output to console/terminal where Streamlit runs
    )

    st.info("CrewAI is starting the presentation generation process...")
    try:
        # Make sure the input to kickoff is a dictionary if needed, or handle appropriately
        inputs = {'topic': topic} # Example if kickoff expects a dict
        result = presentation_crew.kickoff(inputs=inputs) # Pass inputs if required by your Crew setup
        st.success("CrewAI finished generating content!")
        return result
    except Exception as e:
        st.error(f"An error occurred while running the CrewAI process: {e}")
        # You might want to log the full traceback here for debugging
        # import traceback
        # st.error(traceback.format_exc())
        return None

# --- Streamlit App UI ---

st.set_page_config(page_title="AI PPT Generator", layout="wide")
st.title("🚀 AI Presentation Generator using CrewAI")
st.markdown("Enter a topic and choose an LLM. API keys must be configured in secrets or environment variables.")

# --- Sidebar for Configuration ---
with st.sidebar:
    st.header("Configuration")

    # LLM Provider Selection
    llm_provider = st.selectbox(
        "Choose LLM Provider",
        ("Groq", "Google Gemini", "OpenAI"), # Add more if needed
        index=0, # Default to Groq
        help="Select the Language Model provider."
    )

    # --- Secure API Key Retrieval ---
    # Keys are retrieved here using the get_api_key function.
    # No input fields are shown to the user.
    api_key = None
    if llm_provider == "Groq":
        api_key = get_api_key("Groq", "GROQ_API_KEY")
        if not api_key:
            st.warning("Groq API Key not found. Please set `GROQ_API_KEY` in your Streamlit secrets or `.env` file.")
    elif llm_provider == "Google Gemini":
        api_key = get_api_key("Google Gemini", "GOOGLE_API_KEY")
        if not api_key:
            st.warning("Google API Key not found. Please set `GOOGLE_API_KEY` in your Streamlit secrets or `.env` file.")
    elif llm_provider == "OpenAI":
        api_key = get_api_key("OpenAI", "OPENAI_API_KEY")
        if not api_key:
            st.warning("OpenAI API Key not found. Please set `OPENAI_API_KEY` in your Streamlit secrets or `.env` file.")

    # --- Secure Serper API Key Retrieval ---
    st.markdown("---")
    st.markdown("**Web Search Configuration**")
    serper_api_key = get_api_key("Serper", "SERPER_API_KEY")
    if not serper_api_key:
        st.warning("Serper API Key not found. Please set `SERPER_API_KEY` in secrets or `.env`. Web search may fail.")
    # else: # Optional: Indicate success
        # st.success("Serper API Key loaded.", icon="✅")


    st.markdown("---")
    st.info("API Keys are loaded automatically from Streamlit secrets (recommended) or a local `.env` file.")


# --- Main Area ---
topic = st.text_input("Enter the presentation topic:", placeholder="e.g., The Future of Renewable Energy")

if st.button("✨ Generate Presentation"):
    # Perform checks before starting
    if not topic:
        st.warning("Please enter a topic.")
    elif not llm_provider:
        st.warning("Please select an LLM provider in the sidebar.") # Should not happen with selectbox default
    elif not api_key:
        # Specific warning already shown in the sidebar by the retrieval logic
        st.error(f"Cannot proceed without the API key for {llm_provider}. Please configure it in secrets or .env.")
    elif not serper_api_key:
        # Specific warning already shown in the sidebar
        st.error("Cannot proceed without the Serper API key for web research. Please configure it in secrets or .env.")
    else:
        # Initialize LLM (only if api_key is valid)
        llm = get_llm(llm_provider, api_key)

        if llm:
            with st.spinner(f"Generating presentation on '{topic}' using {llm_provider}... This may take a few minutes."):
                # Run CrewAI (only if serper_api_key is also valid)
                crew_result = setup_and_run_crew(topic, llm, serper_api_key)

                if crew_result:
                    st.subheader("Raw CrewAI Output:")
                    # Ensure result is displayed as string
                    crew_output_str = str(crew_result)
                    st.text_area("Content Generated", crew_output_str, height=200)

                    # Parsing logic remains the same
                    parsed_slides = parse_crew_output(crew_output_str)

                    st.subheader("Parsed Slide Structure:")
                    if parsed_slides:
                        st.json(parsed_slides) # Display the parsed structure for verification

                        # Generate PPT
                        st.info("Generating PowerPoint file...")
                        try:
                            ppt_file = create_ppt(parsed_slides, topic)

                            st.subheader("Download Presentation")
                            st.download_button(
                                label="Download PowerPoint (.pptx)",
                                data=ppt_file,
                                file_name=f"{topic.replace(' ', '_').lower()}_presentation.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            )
                            st.success("PowerPoint file generated successfully!")
                        except Exception as e:
                            st.error(f"Error generating PowerPoint file: {e}")
                            # import traceback
                            # st.error(traceback.format_exc())
                    else:
                        st.warning("Could not parse the generated content into slides. Please check the raw output.")
                else:
                    st.error("Failed to get results from CrewAI. Check console logs for details.")
        else:
             st.error(f"Failed to initialize the {llm_provider} LLM. Cannot generate presentation.")

