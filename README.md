# AI-Powered B2B Lead Intelligence & Qualification Agent

An autonomous, multi-stage AI agent that discovers, scrapes, and qualifies B2B leads using search engines, social media (LinkedIn/Xing), and Large Language Models (LLMs).

Unlike basic web scrapers, this system uses an LLM to read website content like a human sales development representative (SDR) — reasoning about services offered, deducing portfolio sizes, identifying decision-makers, and extracting direct personal emails without hallucinating.

## Features
* **Multi-Stage Discovery:** Uses Google Search (via SerpApi) to find direct company sites, and cross-references LinkedIn and Xing to discover companies that lack strong SEO.
* **Smart Scraping:** Bypasses cookie banners and actively hunts down `Impressum` and `Team` pages to find legal representatives.
* **LLM-Powered Extraction:** Uses Cohere (`command-r-08-2024`) with strict Pydantic JSON schemas to extract verified data.
* **Deterministic Scoring:** Grades each lead on a custom 13-point KPI rubric to separate high-value targets (Tier A) from low-value ones (Tier C).
* **Data Minimization:** Only extracts publicly available business data (GDPR compliant for B2B legitimate interest) and explicitly validates emails against source text.
* **Executive Deliverable:** Automatically generates a formatted 4-tab Excel workbook (Overview, Leads, KPI & Method, Sources & QA).

## Setup
1. Clone the repository.
2. Install dependencies: `pip install cohere serpapi pandas openpyxl beautifulsoup4 requests pydantic`
3. Rename `secrets.toml.example` to `secrets.toml` and add your API keys.
4. Run `python b2b_lead_agent.py`.

## Customizing for Your Industry
This template is currently configured with a working example targeting **Property Management Firms in Berlin**. To adapt it to your specific B2B target market, update the following in `b2b_lead_agent.py`:
1. `GENERAL_QUERIES` and `SOCIAL_QUERIES`: Change the search terms to your target industry and region.
2. `DetailedLeadExtraction` (Pydantic Schema): Update the boolean flags and variables to match your ideal customer profile (ICP).
3. `calculate_lead_score()`: Adjust the point values to match the buying signals relevant to your product.
