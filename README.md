# Tony's Excel AI Assistant

An AI-powered Excel add-in that helps with formulas, data cleaning, and analysis using OpenAI GPT-4.

## Features

### 🤖 Formula Generator
- Describe what you want in plain English
- AI generates the Excel formula
- One-click insert into selected cell

**Examples:**
- "Calculate 10% bonus if sales > 1000" → `=IF(B2>1000, B2*0.1, 0)`
- "Count unique values in column A" → `=SUMPRODUCT(1/COUNTIF(A:A,A:A))`
- "VLOOKUP with approximate match" → `=VLOOKUP(A2, B:C, 2, TRUE)`

### 🧹 Data Cleaner
- Analyzes selected data for issues
- Finds typos, inconsistencies, missing values
- Suggests fixes

### 📖 Formula Explainer
- Paste any Excel formula
- Get plain English explanation
- Learn complex formulas

## Installation

### Method 1: Sideload (Developer)
1. Download this repository
2. Open Excel (desktop or web)
3. Go to Insert → Add-ins → Upload My Add-in
4. Select `manifest.xml`
5. The AI Assistant tab will appear

### Method 2: AppSource (Coming Soon)
- Search "Tony's Excel AI Assistant" in Office Store
- One-click install

## Setup

1. Get an OpenAI API key from [platform.openai.com](https://platform.openai.com)
2. Open the AI Assistant panel in Excel
3. Enter your API key (stored locally)
4. Start using AI features!

## Usage

### Generate Formula
1. Select the cell where you want the formula
2. Open AI Assistant panel
3. Click "Formula" tab
4. Describe what you need
5. Click "Generate Formula"
6. Click "Insert into Cell"

### Clean Data
1. Select the data range
2. Click "Clean Data" tab
3. Click "Analyze Data"
4. Review AI suggestions

## Technology Stack

- **Frontend:** HTML, CSS, JavaScript
- **Office Integration:** Office.js API
- **AI:** OpenAI GPT-4o-mini API
- **Platform:** Excel Desktop & Web

## Privacy & Security

- ✅ API key stored locally in browser
- ✅ Data sent directly to OpenAI (not our servers)
- ✅ No data persistence
- ✅ SOC 2 compliant via OpenAI

## Cost

- Uses your own OpenAI API key
- GPT-4o-mini: ~$0.0006 per 1K tokens
- Average request: ~$0.01-0.05
- Much cheaper than enterprise solutions

## Why Not Just Use ChatGPT?

| Feature | Tony's Excel AI | ChatGPT |
|---------|-----------------|---------|
| **Integration** | One-click insert | Copy/paste |
| **Context** | Knows selected cells | No Excel context |
| **Workflow** | Stays in Excel | Switch apps |
| **Cost** | Pay per use | $20/month |
| **Privacy** | Your API key | OpenAI's system |

## Author

**Tony Tan**
- Tax Technology & Transformation at EY
- Product Manager & AI Developer
- [LinkedIn](https://www.linkedin.com/in/tonytan3/)
- [GitHub](https://github.com/TonyTan12)

## License

MIT License - Feel free to use and modify!

## Contributing

Pull requests welcome! This is an open-source project.

## Support

- Open an issue on GitHub
- Email: Tonytan1999aol@gmail.com

---

**Built with ❤️ for Excel power users everywhere**