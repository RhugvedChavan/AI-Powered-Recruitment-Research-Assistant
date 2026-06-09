from langchain_mistralai import ChatMistralAI
from tavily import TavilyClient
from dotenv import load_dotenv
import os


# LOAD ENVIRONMENT VARIABLES

load_dotenv()

MISTRAL_API_KEY = os.getenv("MISTRAL_API_KEY")
TAVILY_API_KEY = os.getenv("TAVILY_API_KEY")

if not MISTRAL_API_KEY:
    raise ValueError("MISTRAL_API_KEY not found in .env")

if not TAVILY_API_KEY:
    raise ValueError("TAVILY_API_KEY not found in .env")


# INITIALIZE CLIENTS


llm = ChatMistralAI(
    model="mistral-large-latest",
    temperature=0.1
)

search_client = TavilyClient(api_key=TAVILY_API_KEY)


# MODE SELECTION


print("\n========== BUSINESS INFORMATION CHATBOT ==========")
print("1. Company")
print("2. College")


choice = input("\nSelect Mode (1/2): ").strip()

mode_map = {
    "1": "Company",
    "2": "College",
}

if choice not in mode_map:
    print("Invalid choice!")
    exit()

mode = mode_map[choice]

name = input(f"\nEnter {mode} Name: ").strip()

# SEARCH QUERY


queries = [
    f"{name} official website",
    f"{name} GSTIN GST status CIN PAN",
    f"{name} directors DIN number",
    f"{name} employee count workforce size",
    f"{name} annual revenue turnover growth",
    f"{name} linkedin company page",
    f"{name} careers jobs hiring",
    f"{name} HR manager talent acquisition manager",
    f"{name} headquarters branch locations",
    f"{name} contact details email phone"
]
print("\nSearching web...")

search_results = search_client.search(
    query=search_query,
    search_depth="advanced",
    max_results=10
)

# COMBINE SEARCH RESULTS

context = ""

for idx, result in enumerate(search_results.get("results", []), start=1):
    context += f"\nSource {idx}\n"
    context += f"Title: {result.get('title', '')}\n"
    context += f"URL: {result.get('url', '')}\n"
    context += f"Content: {result.get('content', '')}\n"
    context += "-" * 80 + "\n"


# PROMPT

prompt = f"""
You are a Business Intelligence Assistant.

Entity Type: {mode}
Entity Name: {name}

Use ONLY the information provided below.

If any field is unavailable write:
Not Found

Return information in this format.

For Company:

Company Name:
Company Name:
Company Overview:
Registered Office Address:
Industry Type:
Sub-Industry:
Headquarters & Branch Locations:
GST Registration Number:
Date of Establishment:
GST Status (Active/Cancelled):
PAN Number:
CIN/UID Number:
Company Constitution (Pvt Ltd / Public Ltd / LLP / Proprietorship):
Names of Directors and their DIN Numbers:
Employee Strength:
Revenue & Business Growth (Last 3 Years if Available):
Official Website:
LinkedIn URL:
Current Hiring Status:
Active Job Openings:
Hiring/Expansion Plans:
HR Manager or Talent Acquisition Manager Name:
Publicly Available Contact Details:
LinkedIn Profile of HR Manager:
Recruitment Activity Level (High/Medium/Low):


For College:

College Name:
Website:
Email:
Contact Number:
Address:
City:
Country:
Student Count:
Faculty Count:
Courses Offered:
Departments:

College Description:


WEB SEARCH DATA:
{context}
"""


# GENERATE ANSWER

response = llm.invoke(prompt)

print("\n" + "=" * 80)
print(response.content)
print("=" * 80)

