# 🚀 AI Presentation Generator

An automated PowerPoint presentation generator built with Streamlit and CrewAI that creates professional slide decks on any topic.

## 📋 Overview

This application uses AI to generate complete PowerPoint presentations on any topic. It leverages:

- **CrewAI**: Orchestrates a team of specialized AI agents that work together
- **LangChain**: Connects to various LLM providers
- **Streamlit**: Provides the web interface
- **python-pptx**: Creates the PowerPoint files

## ✨ Features

- Generate complete presentations with just a topic
- Choose from multiple LLM providers (Groq, Google Gemini, OpenAI)
- Automatic web research using Serper API
- Downloadable PowerPoint (.pptx) files
- Secure API key management

## 🛠️ How It Works

1. **Research Agent**: Searches the web for information on your topic
2. **Outline Agent**: Creates a logical presentation structure
3. **Content Agent**: Writes bullet points for each slide
4. **PowerPoint Generation**: Converts the content into a downloadable .pptx file

## 🔧 Setup & Installation

### Prerequisites

- Python 3.10+
- API keys for:
  - An LLM provider (Groq, Google Gemini, or OpenAI)
  - Serper.dev (for web search capabilities)

### Installation

1. Clone the repository:
   ```
   git clone <[repository-url](https://github.com/piyush230502/Autonomous_PPT_Agent.git)>
   cd ai-presentation-generator
   ```

2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Set up API keys:
   - For local development: Create a `.env` file in the project root
   - For deployment: Add to Streamlit secrets

   Example `.env` file:
   ```
   GROQ_API_KEY=your-groq-api-key
   GOOGLE_API_KEY=your-google-api-key
   OPENAI_API_KEY=your-openai-api-key
   SERPER_API_KEY=your-serper-api-key
   ```

   Example `.streamlit/secrets.toml`:
   ```
   GROQ_API_KEY = "your-groq-api-key"
   GOOGLE_API_KEY = "your-google-api-key"
   OPENAI_API_KEY = "your-openai-api-key"
   SERPER_API_KEY = "your-serper-api-key"
   ```

## 🚀 Usage

1. Run the application:
   ```
   streamlit run app.py
   ```

2. In the web interface:
   - Enter your presentation topic
   - Select your preferred LLM provider
   - Click "Generate Presentation"
   - Wait for the AI to create your presentation (typically 2-5 minutes)
   - Download the PowerPoint file

## 🧩 Project Structure

```
.
├── .streamlit/
│   └── secrets.toml  # Store API keys securely for deployment
├── app.py            # Main Streamlit application code
├── requirements.txt  # Python dependencies
└── README.md         # This file
```

## 🔄 How the Code Works

1. **API Key Management**: Keys are securely retrieved from Streamlit secrets or environment variables
2. **LLM Selection**: The application supports multiple LLM providers through LangChain
3. **CrewAI Setup**: Three specialized agents work together in sequence:
   - Researcher: Gathers information using web search
   - Outliner: Creates a logical presentation structure
   - Writer: Generates bullet points for each slide
4. **Content Parsing**: The AI output is parsed into a structured format
5. **PowerPoint Generation**: The python-pptx library creates slides with proper formatting

## 📝 Notes

- The quality of presentations depends on the LLM used
- Groq with Llama 3 models offers a good balance of speed and quality
- For more complex topics, consider using more powerful models like GPT-4o

## 📄 License

[MIT License](LICENSE)
