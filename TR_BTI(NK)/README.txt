INPUT:
      1.TR_BTI_Template.xlsx
      2.TR_BTI.xlsx
      3.req_file.xlsx

NOTE:
Step1: connect The Chicago vpn and run "BTI_DBData_Download(1).R" it will download the records from DB.
step2: disconnect VPN and Run the script "BTI_Records_Download(2).R" it will extract the data from Website.
step3: Run the "BTI_Records_Validation(3).R" this script it will validate the records and it will give Adition records in text file "bti.txt"
step4: Run the "BTI_Rulings_Download(4).R" it will download the rulings for Adition records.
step5:verify the "TR_BTI_Update(10Jul2023).xlsx" file wherever Image coloumn contain "N" verify with source it is genuine or not.

step6:######### TURKEY BTI PDF CREATION INSTRUCTION DOCUMENT ########

-> "Turkey_BTI.R" This tool run time depends on the number of records of input file.
-> First we need to place updated turkey BTI extraction file with the fixed format in IN_FILES folder.
-> The file name should be "TR_BTI"
-> Before running the tool need to delete all files in "excel_files" folder and "images" folder.
-> This tool will create excel files.
-> We need to convert these excel files to pdf.
-> For convert these excel files to pdf, we need to take these files to our system and
   need to run UIpath tool.



