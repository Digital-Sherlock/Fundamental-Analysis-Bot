# Fundamental-Analysis-Simplified

## The Problem Statement

In the common era of information economy people have access to endless amount of resources and it's easy to get lost. In some cases this is the reason why some people may not even start learning a new venture, explore a subject, switch their careers, etc.

Similar applies to investing with huge amounts platforms, resources, insights, videos available in the Internet. Many platforms that offer public stock information (i.e. news, financial statements) cat get overwhelming for an average investor who wants to make a basic evaluation of a company as a potential investment. In addition, some of these resources introduce additional cost which might not be ideal for everyone.

## The Goal of the Project

The goal of this project is simplify fundamental analysis of a stock and bring everything in one place in a form of a familar Excel spreadsheet.

## The Outcome

Here's an example of how the end result will look like. The screenshot below as a sample of an Income Statement of a stock with randomly generated financials. The final Excel file will contain three sheets: Income Statement, Balance Sheet, and Cash Flow Statement.
![image](https://github.com/Digital-Sherlock/Fundamental-Analysis-Bot/assets/66618495/60a69111-6879-40f4-9281-4170f094ec86)

## Details

The code and specifically the stock data is based on the financial IEX Cloud API. This particular project used Sandbox API endpoints.

## Issues

The main issue is a lack of some stock information that sandbox IEX Cloud API is limited on. To get a full and complete data one will have to switch to a paid version of IEX Cloud.

## Code

The code is organized in four sections. The naming convention of each piece aims to make it clear what part of the stock's fundamentals it covers. The Excel_Formatter brings all sections and creates an Excel file.