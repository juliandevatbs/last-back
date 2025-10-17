def read_punctual_sheet_data(workbook, sheet_name: str) -> {}:



    try:

        punctual_sheet = workbook[sheet_name]

        hours = {



        }



        hours["1"] = punctual_sheet["F71"].value
        hours["2"] = punctual_sheet["R71"].value
        hours["3"] = punctual_sheet["F71"].value


        return hours


    except KeyError:

        print(f"Sheet {sheet_name} not found please review")
        return {}

    except Exception as e:

        print("Error opening the sheet")
        return {}

