def aggregate_xlsheet(dpath):
    
    ''' 
    a function to aggregate excel sheets from one or several workbooks into one excel file (workbook).
    
    Parameter
    =========
    
    dpath: full path string of the directory where the files to be aggregared are kept.
    
    '''
    
    try:
    
        # define the path where the excel files to be aggregated are kept
        source_filepath = Path("/Users/xxxx/Downloads/xltest")

        # change the path to become the current working directory
        os.chdir(source_filepath)

        # create a new folder in the current working directory
        Path('merged').mkdir()

        # insert a new workbook in the folder
        nwbk = pyxl.Workbook()
        
        # name the new workbook as "merged.xlsx" and define its path 
        merged_wbk = source_filepath / 'merged' / "merged.xlsx"
    
        # iterate through each file in the source folder
        for file in tqdm(os.listdir(source_filepath)):

            # select only excel file 
            if file.endswith(".xlsx"):

                # load the excel file and set it active
                wbk = pyxl.load_workbook(source_filepath / file)
                wbk.active

                # iterate through each sheet in the workbook
                for sh in wbk.worksheets:

                    # for each selected sheet in the current workbook, create a new sheet in the destination workbook (file)
                    nwbk.active
                    nsh = nwbk.create_sheet(sh.title)

                    # iterate through the rows and cells in the selected sheet and write data into the new sheet
                    for row in sh:
                        for cell in row:
                            nsh[cell.coordinate].value = cell.value

                    # save the new workbook
                    nwbk.save(merged_wbk) 

        # load the new workbook
        bk = pyxl.load_workbook(merged_wbk)

        # iterate through the sheets and remove any sheet without data
        for sh in tqdm(bk.worksheets):
            if sh.max_row == 1 and sh.max_column == 1:
                bk.remove(sh)

        # save the workbook
        bk.save(merged_wbk)
        
        print("Sheets aggregation was successful!")
        
    except Exception as err:
        print(err)

if __name__=='__main__':
    aggregate_xlsheet("/Users/xxxx/Downloads/xltest")   