# Finance Analysis VR Bank

## Description

This script should help you making your financial overview simpler. It takes csv Exports of your income and expenses and condeses them in minimalistic diagrams and overview for you to easier cross-check your financial decisions of the past.

## Limitations

This script is optimized for the csv structure of the Volksbanken Raiffeisenbanken Ludwigsburg and may not work out of the box with your csv file.

## Usage

### Preparation

Start by creating a folder called `Bank Exports` in which you download the csv exports. Name them as you like, the program will scan all contents of this folder. I recommend YYYY.MM.csv. Choose wisely as this will also be the sheet name in the Excel later on and the sorting will happen according to the name alphabetically.

Next, create the `categories.json` file for categorising your expenses. Use the empty structure shown below:

```
{
    "categories": [
    ],
    "receiver": {
    },
    "reason": {
    }
}
```

It will get filled via prompts while running the program.

For the last Step, create an empty 1 Sheet Excel file and name it `Analyse.xlsx`. Here all the data will be dumped afterwards.

### Action

You may now run the analyse.py file.

It will automatically categorise your expenses and income based on the `categories.json` contents. Because it is empty at the first time, it will prompt you for every payment.

Follow the instructions and create categories by typing the name or reuse them based on the shown index (type the number). Tell the script if it should automatically assign the category based on the receiver or the reason or if it should not use similar information for further categorising. You can also specify if some other keyword should be used in the future. (e.g. shorten "AMAZON EU S.A R.L., NIEDERLASSUNG DEUTSCHLAND" to "Amazon") Th Script is not case senstive.

When the program finishes, you will have an empty sheet whcih was used to create the file. After that a sheet with the overview of all the data you have analysed. after that you will have a sheet for every month (or timeframe you exportet the data for) with the analysis for this month.

# Future

If you are interested in the expansion of this script hit me up then we can sort out how we can make it happen.

Things I have in mind right now but not enough time at hand:

- "Legacy" support for PDF format of 'Kontoauszüge'
- Maybe the charts a little more pretty
- Possibility to ignore payments in the chart/analysis (e.g. to family members, saving accounts etc.)