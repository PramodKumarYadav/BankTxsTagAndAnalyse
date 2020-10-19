# BankTxsTagAndAnalyse
# Goal: 
I created this macro to analyse my quarterly bank statements to filter out txs relevant for my company. 
Note: This is an ugly solution that works. I could have done this better but I have worked in VBA long back and didnt really want to spend more time working on old tech (this macro). If you are interested to improve, you are more then welcome to create a PR and I will be happy to merge it in the master. 

# Note:
This is an early version and if you decide to try, you will try it at your own risk. If you find any bugs, feel free to report or fix yourselves in a PR. I do not gurantee that this is error free (since I havent tested it enough). I hope it works okay but its free of any legal obligations (basically do not sue me if it doesnt work for you :) - try at  your own risk). 

1. Download your ABN AMRO bank txs in excel format in an XLS.
2. Save this XLS as XLSM.
3. Rename the txs sheet as 'test'
4. Create a new sheet and rename as 'Tags'
5. Add your Tags in column B (starting first row). Say as:
Albert Heijn
Aldi
Amazon Payments
ANWB B.V.
Apotheek
ARTSEN ZONDER GRENZEN
Blokker bv
H&M
HEMA
Intertoys
Jumbo
MCDONALD
OV-chipkaart
(add more relevant to you)
6. Copy the VBS macro from here.
7. Open code editor in XlSM by pressing ALT+F11 (or going to developer tab and pressing Visual basic button)
8. Create a new module and copy paste the code from VBS there.
8. Run the macro by pressing the run button. 
