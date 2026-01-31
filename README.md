# Tutorial ChiquitaAI
*2026 January tutorial* 

Use free AI chatbots to build an Excel for Mac worksheet to manage sharing and owning a private plane.

# Coding Example using ChatGPT

Since 2022, when AI-powered chatbots started being very popular, I experimented how to use their free tier of service to describe Machu Picchu so that the chatbots can generate code for it. The purpose is to help humanitarians, without much coding experience, build state-of-the-art apps to support their daily work.

_In project Machu Picchu, we aim to use 21st century tools to change the way humanitarian aid is practiced._

This document is a step by step illustration of the chatbot-powered coding workflow. To make it easy for laypersons (non-programmers) to reproduce, instead of coding on JavaScript, Python or a programming language, in this example we develop a complete Excel application using Excel best practices with the assistance of the free tier of ChatGPT, in Thinking mode when required.


# Overview of our purpose
- [Start with section Full scale example](#full-scale-example-excel-tool-for-expenses-tracking)
- [Hop to Session 1: Generate Detailed Design using ChatGPT](#session-1-generate-detailed-design)

## AI-powered Workflow in a Nutshell
![overview](./images/000_overview.png)
- Step A: 1️⃣ describe what you want 2️⃣ have the bot criticize and suggest improvements 3️⃣ have it generate a detailed design. Reiterate once, to trap eventual hallucinations.
- Step B: start a new session, 1️⃣ paste the design, 2️⃣ have the AI chatbot generate the code, 3️⃣ have it criticize its code. Reiterate once, to trap eventual hallucinations.
- Step C: 1️⃣ run the AI-generated app, 2️⃣ handcraft add the desired features that are too business-specific to spend time explain to the chatbot, 3️⃣ start a new session , submit my code, ask for refactoring
- Loop to step B.

Hints:

- Instead of describing textually what you want, you make a quick prototype and give it to the chatbot to explore and comment.
- Often, irrelevant criticisms of the chatbot about your prototype announces future misunderstandings when asking it to generate code. Clarify these criticisms and start a new session.
- Start new session as often as practical, as soon as the chatbot has achieved a significant step. This will clean the context, focus the discussion and avoid hallucinations.

Illustration:
- Step A.1 is illustrated in "No-Think" mode by prompt 1 [(./prompt_01.md)](./prompt_01.md), the answer of ChatGPT is [(answer_01.md)](./answer_01.md)
- Step A.1 is illustrated in "Think" mode by the same prompt 1 [(prompt_01.md)](./prompt_01.md), the answer of ChatGPT is [(answer_01b.md)](./answer_01b.md)
- Step A.2 is a question for example by prompt 2 [(prompt_02.md)](./prompt_02.md), the clarification of ChatGPT is [(answer_02.md)](./answer_02.md)
- Step A.3 is illustrated in "Think" mode by prompt 3 [(prompt_03.md)](./prompt_03.md), the clarification of ChatGPT is [(answer_03.md)](./answer_03.md)

## Spoiler Results

Spoiler results list what we have learned from this tutorial if you don't have time to go into details:

- ChatGPT 5.2 hallucinates the least when the question is difficult and not widely known in public documents. For example, when on Excel for Mac a I use a LAMBDA function inside a MAP function to populate 2D array cells created with MAKEARRAY.
  -  Don't be scared by these names, I barely know them too before I started this project, but I learned how to use them with the chatbots.
- Gemini is the second best chatbot. Most of the time it gives pertinent advice. But in some implementation imperfections of Excel on Mac, it keeps reasoning around without reaching a solution. I use Gemini as a cross-check of ChatGPT answer: I give it the code of ChatGPT and ask it to comment and improve.
- Claude Sonnet 4.5 is very popular among programmers but its free tier is not usable for our purposes. It is complicated to set up Claude to upload Excel files, for it to analyze and criticize. This uploading is critical, because we ask the chatbot to comment an implementation.
- Mistral, DeepSeek 3, Qwen 2 (in their free tiers) are truly not up to the task.

So we use ChatGPT 5.2 instead of other chatbots: in this particular case, ChatGPT free tier gave the most useful answers.

Keep in mind that a modern chatbot doesn't simply call on a LLM (Large Language Model) to generate answers. A state-of-the-art chatbot is a complex orchestrator of tools other than the LLM (path generator and evaluator, web crawler, adversarial challenger etc.) to find the best answer. Therefore, an outstanding LLM may perform poorly if it doesn't have well configured its environment of support tools.

_The situation of chatbots using LLMs is similar to engineering a car by starting with a powerful engine and then add supporting accessories to make it a comfortable limousine, or SUV, or a van. In our experience, Claude may be good for coding apps, but less effective to make a real-life Excel worksheet._  

    I haven't tested Claude for Excel because it is a feature of a paid service.

## What this article is made of

This article carries several types of content:
-	The textual explanation is hosted in a Medium article, a LinkedIn article, and a Github README.md file [(https://github.com/kvutien/tutorial_chiquitaAI/)]. It describes our purposes, the chatbot prompts and parts of the chatbot reply to introduce the next prompt. It contains also our comments on the bot replies.
-	The full chatbot answers and the generated codes are in the above github repository.
-	The intermediate worksheets are also in the same github repository.

If you don't have time to look at details, you can stop reading here. If you continue, be prepared to reproduce by yourself the results with a free ChatGPT account and Excel 365.

# Full scale example: Excel Tool for Expenses Tracking
- [Hop to Session 1: Generate Detailed Design using ChatGPT](#session-1-generate-detailed-design)

In this example of January 2026, I use ChatGPT 5.2 in Thinking mode. The 2 other chatbots that have this reasoning mode are : Google Gemini 3 and Claude 4.5, but their reasoning is more shallow.

I want to realize a cost analysis tracking tool for a private aircraft. This is not trivial example project and is composed of 3 sub-projects :

1.	Re-engineer the initial hack of a yearly expense tracking tool. This is the purpose of this post.

2.	After the AI-refactored tool is validated, add the operations costs (fuel, landing fees) to make it a flight billing and operations. This step involves understanding the nature of the activities and is where AI helps a lot because it converts the human language to computer processes.

3.	Re-engineer both the expense tracking, the flight billing and the refueling operations to make statistics and a KPI dashboard. This stage will only be sketched.

To start, I made a sample Excel worksheet to show what I want to do. Most of the formulas are hard coded cell by cell, which is a typical hack, made to describe and explore a solution. 

This tool has 3 Excel tabs:

- "Data": this table carries the bulk expenses data, each expense is assigned to a category

![Screenshot of "Data"](./images/001_proto_screenshot.jpg)

- "Stats": this table carries the yearly expenses per category. It distinguishes the fixed recurring expenses, the variable expenses, and for the exceptional high-value non-recurring expenses that are paid one year and amortized over several years.

![](./images/002_proto_screenshot.jpg)

- "KPI": this table specifies the categories of expenses. As a typical hack, it is a mix of read-only category labels with a textual explanation of each category and of formula cells that calculate the average expenses per year for each category as well as the maintenance performance (ratio of mechanics hours on flying hours).

![](./images/003_proto_screenshot.jpg)

# Session 1: Generate Detailed Design

## Prompt: Describe Data and Actions

I upload this Excel example to Chat GPT (no Thinking, no Deep research) and asked. This prompt is also in the github as  [(`prompt_01.md`)](./prompt_01.md), the answer of ChatGPT is [(answer_01.md)](./answer_01.md):

    Keep in mind for the whole session that you are a highly experienced Excel programmer, you use Excel 365 for Mac with EU locales and, although I'm familiar with Excel, there are many advanced aspects I'm not too familiar with. You know that Power Pivot and Data Model are not implemented in Excel for Mac.

    Describe in detail the uploaded file so that you can use your description as requirements to create an Excel worksheet with the same functionalities but better engineered and using Excel best practices, with easy-to-understand variables and comments.

The answer of ChatGPT follows. This answer is also in the github as [(`answer_01.md`)](./answer_01.md)

## ChatGPT Describes Data and Actions ("No Thinking" mode)
### My comments on ChatGPT no-thinking answer 
The overall structure is correctly perceived.
- Overall Purpose of the Workbook
- Sheet: OLD-Data (Raw Data Layer)
- Sheet: OLD-Stats (Yearly Aggregation & Cost Model)
- Sheet: OLD-KPI (High-Level Indicators)
- Cross-Sheet Dependencies
- Target Architecture (Implicit Requirements

The overall purpose and the layers of sheets are identified. The major weaknesses of the implementation are listed.

Sheet "Data" is well identified as an "almost-Excel-table". Categories are wrongly qualified as "free-text": the data validation is not identified, certainly because we are in no-thinking mode. We have a target functional requirements for re-factoring.

Sheet "Stats" is well qualified as "a manual pivot table implemented with formulas", "extremely layout-dependent, logic and presentation are tightly coupled". "Old-style Excel calculations". Improvements are proposed to reduce the weakness of the programmation style.

Sheet "KPI"is weel perceived as "summarizing KPI" stats across years and very fragile dependent on years count.
Overall, the "no-thinking" mode notes the biggest flaws:  there is no abstraction layer and the same logic is duplicated among similar cells of an matrix, making the sheet very fragile to maintain. However, the Name Manager has not been examined, nor the data validation.

Next, we activate the "Thinking" mode and ask the same prompt. 

##	ChatGPT Describes Data and Actions ("Thinking" mode)
### My comments on ChatGPT thinking answer 
The answer of ChatGPT is in the github as [(`answer_01b.md`)](./answer_01b.md)

The "Thinking" answer is more detailed. This may lure you to accept ChatGPT proposal to start implementing. However, keep in mind that ChatGPT is not a magician who knows exactly what you intend to do. Implementing at this stage of understanding will certainly miss the spot. So I ask for clarification on the proposed solution.

##	ChatGPT clarifies how it'd do statistics without pivot tables ("Thinking" mode)
This asking for clarification is also in the github as  [(`prompt_02.md`)](./prompt_02.md), the answer of ChatGPT is [(answer_02.md)](./answer_02.md):

    Hum. I've asked you to remember during the whole session that you are using EU locale. Keep in mind also that you are using Excel for Mac with all its limitation (see initial prompt). 
    Now show an example of your sheet Calc_Yearly once filled.

The answer of ChatGPT follows. 

### My comments on ChatGPT thinking clarification 
This answer is also in the github as [(`answer_02.md`)](./answer_02.md)

The proposition of ChatGPT is to organize the years by rows and the calculated values as columns. The rows will grow as more years are added while the calculated values will very rarely have new members. Then we can write one formula per value and let Excel still down the formulas.

Such formulas are much more fexible than ppredefined pivot components. Furtherore, the calculation logic is not hidden in pivots.

##	ChatGPT summarizes what it understood and assumptions to confirm

Again, ChatGPT concludes its aswer by proposing to code. But it is not yet a good idea.
-	To explain precisely your intentions, either you write a very detailed prompt, or you illustrate (like here) with a quick Excel prototype that you ask ChatGPT to analyze and criticize.
-	If you illustrate by a prototype, it gives ChatGPT a rough idea of what you want. If you ask it to code immediately after this early analysis of the prototype, ChatGPT will fill the blanks with its own training memories. If you disagree with the output, you can correct it but unless you say explicitly that its assumptions are wrong, ChatGPT will keep polluting the context with them.
-	At this stage, instead of coding, I recommend to clear up with what ChatGPT has assumed wrongly. From these corrections, ask it which assumptions that your latest corrections made it modify.

My corrections are in the prompt stored in prompt 3 [(prompt_03.md)](./prompt_03.md). The clarification of ChatGPT is answer 3 [(answer_03.md)](./answer_03.md):

-	*The purpose of re-factored worksheet will be used to calculate the balance of each pilot, based on the general cost of maintaining the plane, the costs of their usage of the plane and their contributions.*
-	*The re-factored worksheet will be composed of the data entry sheets and read-only sheets.*
-	*Data entry sheets are composed of Expenses, AVGAS (fueling), Flights (pilot, time, location, Hobbs, etc.) and Contribs. All data entry sheets have a column FundSource to trace who is paying.*
-	*Read-only sheets are composed of Params, Stats, Overview (Dashboard & KPI) and Audits. A last sheet Misc is for temporary calculations and a README sheet is the user guide.*
-	*Sheet Flights has data entry columns and helper columns to calculate costs related to the trip of the row. These helper columns will be read by sheet Stats to calculate the balances.*
-	*Sheet Expenses (Data) needs no Subcategories. But it will need a helper column Type (Fixed, Variable, Exceptional).*
-	*Sheet Expenses (Data) will never contain flight and Hobbs value, nor AVGAS. This information will be in a distinct sheet Flights, respectively AVGAS.*
-	*Sheet AVGAS has data entry columns and a few helper columns that will be used by Stats to calculate fuel data.*
-	*Sheet Expenses (Data): In Europe there are a lot of VAT rates depending on the year and on the country. Defining TVARate by XLOOKUP is overly complicated.*
-	*Sheet Params (Lookups): There are some parameters that can be used with XLOOKUP like the dry hourly cost for pilots. Some others such as fuel price, consumption per hour, fuel per hour, fuel cost per hour, will be the average since the beginning until the moment of the flight.*
-	*Sheet Params (Lookups): Ignore. There is currently no amortization yet. Only in future versions.*
-	*Sheet Stats (Calc_Yearly): Very interesting use of SUMIFS which improves on pivot tables. We keep this solution from now on. We will call this sheet Stats.*
-	*Sheet Stats (Calc_Yearly): Keep in mind that the SUMIFS giving average hourly consumption, fuel price, yearly variable costs, fixed costs fuel costs are reused to calculate the cost of each trip in Flights. We must avoid calculation loops.*
-	*Sheet Stats (Calc_Yearly): Let's ignore amortization for the time being. It will be introduced later after a few months using this Excel worksheet.*
-	*Sheets KPIs & Dashboard: Merge these 2 into a sheet named Overview. It is the read-only surface while Stats is the calculation engine.*
-	*Sheet Audits is a good contribution that ChatGPT "think mode" brought in the last answer.*

*Acknowledge that you'll take these clarifications into account. Ask me to confirm assumptions if you still have that I have not clarified. Then in the coming prompts I'll give you details on how to calculate costs in the calculation formulas.*

As you may see, I add now precisely clarifications to what ChatGPT has in its memory. This kind of prompt is very close to fine-tuning a model, but limited to a session.

The answer of ChatGPT is in the github as [(`answer_03.md`)](./answer_03.md)

### My comments on ChatGPT (thinking) summary of progress 

# a