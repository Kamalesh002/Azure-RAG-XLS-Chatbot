RAG Q&A Chatbot with Excel Upload
This project implements a Retrieve and Generate (RAG) Q&A chatbot using Azure OpenAI and Azure Cognitive Search, allowing users to interact with data stored in Excel files. Users can upload an Excel file, and the chatbot will index the content, enabling users to ask questions related to the data. The chatbot retrieves the relevant information from the indexed documents and generates responses using GPT-style models.

Features
Excel File Upload: Users can upload .xls or .xlsx files, which are indexed and stored in Azure Cognitive Search.

AI-Powered Q&A: Leverages Azure OpenAI to process and generate responses to user questions based on the indexed data.

Searchable Data: All content from the uploaded Excel files is indexed for fast retrieval, allowing users to ask context-based questions.

Contextual Answer Generation: Uses the context from the indexed content and OpenAI’s GPT-4 model to generate accurate answers.

Tech Stack
Backend: Python, Flask

Frontend: HTML, CSS (for chat interface)

AI: Azure OpenAI (GPT-4), Azure Cognitive Search

File Handling: Pandas, openpyxl for handling Excel files

Cloud Services: Azure (for OpenAI, Cognitive Search, App Insights)

Logging & Monitoring: Application Insights

Prerequisites
Before running the project, make sure you have:

An Azure account with access to Azure OpenAI and Azure Cognitive Search.

A Python 3.8+ environment.

pip to install Python dependencies.

Setup
Step 1: Clone the repository
bash
Copy
Edit
git clone https://github.com/<your-username>/rag-qa-chatbot.git
cd rag-qa-chatbot
Step 2: Install the dependencies
Create a virtual environment and install the required dependencies:

bash
Copy
Edit
python3 -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`
pip install -r requirements.txt
Step 3: Configure environment variables
Create a .env file in the root of the project and add the following Azure credentials (You can find these in your Azure portal):

ini
Copy
Edit
AZURE_SEARCH_ENDPOINT=your_azure_search_endpoint
AZURE_SEARCH_KEY=your_azure_search_key
AZURE_SEARCH_INDEX=your_search_index_name

AZURE_OPENAI_ENDPOINT=your_azure_openai_endpoint
AZURE_OPENAI_KEY=your_azure_openai_key
AZURE_OPENAI_DEPLOYMENT=your_openai_deployment_name
AZURE_OPENAI_EMBED_DEPLOYMENT=your_embedding_deployment_name

APPINSIGHTS_INSTRUMENTATIONKEY=your_application_insights_instrumentation_key
Step 4: Run the application
bash
Copy
Edit
python app.py
This will start a Flask web server on http://127.0.0.1:5000/.

Usage
1. Upload an Excel File
Navigate to the web interface (default: http://127.0.0.1:5000/). Use the file upload section to select and upload your Excel file.

2. Ask Questions
Once the file is uploaded, you can ask questions related to the data in the file. The chatbot will retrieve the relevant context from the indexed data and generate responses using the Azure OpenAI model.

Directory Structure
bash
Copy
Edit
├── app.py                 # Flask application
├── rag_utils.py            # Utility functions for embeddings and indexing
├── templates/              # HTML templates
│   └── chat.html           # Frontend UI for chatbot
├── .env                    # Configuration file for environment variables
├── requirements.txt        # List of dependencies
└── uploads/                # Folder for storing uploaded Excel files
How It Works
Upload Excel File: The user uploads an Excel file containing data.

Indexing: The content of the Excel file is indexed using Azure Cognitive Search. The text is embedded and stored, enabling fast retrieval.

Search & Chat: The user can ask questions, and the chatbot generates responses based on the indexed content using Azure OpenAI’s GPT model.

Logging: All logs, including application events and errors, are tracked using Azure Application Insights.

Configuration Options
File Extensions: Only .xls and .xlsx file formats are supported for upload.

Search Index: The content of the uploaded Excel file is indexed into Azure Cognitive Search with embeddings for efficient querying.

AI Model: The OpenAI GPT model is used for generating answers based on the search results.

Logging & Monitoring
The application integrates with Azure Application Insights for detailed monitoring and logging. All interactions with the chatbot, errors, and performance metrics are logged to the Azure portal for further analysis.

Contributing
We welcome contributions to improve the chatbot! If you'd like to contribute, please follow these steps:

Fork the repository.

Create a new branch (git checkout -b feature-name).

Make your changes and commit (git commit -am 'Add new feature').

Push to your branch (git push origin feature-name).

Open a pull request.

License
This project is licensed under the MIT License - see the LICENSE file for details.

Acknowledgments
Azure OpenAI for providing GPT-4 API for contextual response generation.

Azure Cognitive Search for enabling powerful search and retrieval capabilities.

Pandas and openpyxl for handling Excel data.

Flask for providing a simple web framework.
