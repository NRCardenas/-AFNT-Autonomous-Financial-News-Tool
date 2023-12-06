# AFNT - Autonomous Financial News Tool
An AI tool (based on LLM) that simplifies the elaboration of financial news reports.

## What is it?

AFNT is a easy to use tool that simplifies the elaboration of daily market reports.
![image](https://github.com/NRCardenas/-AFNT-Autonomous-Financial-News-Tool/assets/153119544/347ba595-9c6e-4bc9-bb87-8dcd6594a75b)
(This is a just an example)

The report include the analysis of the recent both economic and market activity (including events such as central bank announcements, M&A, economic outlooks, presentating of financial reports...), the evolution of various Stock Index, the list of macro economic variables that will (or have been) pusblished during the day and a detailed evolution of the IBEX Index.

This reports is done 3 times each day: at the market opening, half market season and market closure. At this moment, this work only for the European market, but will be include more options in the future.

## How do this work?

![image](https://github.com/NRCardenas/-AFNT-Autonomous-Financial-News-Tool/assets/153119544/603bdda5-9732-4caf-8ae3-45f6f0093fe0)

Breafly, the tool allows the user to create a calendar for the next week macro economic economic events. In the following week, for each day the program let the user introduce the published data for each variable (if it has already been published). The rest of the work is done automaticaly just by running the program. 

The first paragraph is done with a trained LLM and an implementation that allow the model to search the internet (both parts requires set the parameters properly (temperature, stop, tokens, first system instructions, embedings, markets, delay...) and increase the funcionality of the trained LLM). The LLM is trained with expert-like reports.

The commodities and market information comes from yfinance, that has little delay on its publication; this does not represent a problem and the delay for each market can be seen in their page. 

Finally, the recollection of macro economic data is done by the user. This could be done with other pay API.

## Cost
This progam relies on Open AI technology, the price per token and model can be seen on its website. Also, the tool uses BING WEB Search API, that has a free option which is more than perfect for this purpose.

## Requirements and Dependencies

The required libraries are:
 - yfinance
 - openai
 - pandas
 - request
 - tkinter (ttk)
 - openpyxl
 - calendar
 - from docx import Document

Altough it is not a library, the tool uses the BING WEB Search API.

## Contact:

For more information, please feel free to contact any of the authors.

## Creators (github):
Sebastian Arturo Saturno Teles -> saturnos3 
Néstor Rafael Cárdenas Castillo -> NRCardenas
