require 'rubyXL' # rubyXL is a gem which allows the parsing, creation, and manipulation of Microsoft Excel (.xlsx/.xlsm) Documents

input_filename = ARGV[0]

workbook = RubyXL::Parser.parse('./' + input_filename)  # Parsing an existing workbook 

paramsheet = workbook[0] # paramsheet refers to the 'Parameters' sheet of the excel file
worksheet = workbook[1] # worksheet refers to the 'Summary' sheet of the excel file

res_book = RubyXL::Workbook.new # create a new workbook for output
res_sheet = res_book[0]         # res_sheet refers to the first sheet of the output excel file

date = paramsheet[0][1] == nil ? '' : paramsheet[0][1].value
h_YR = paramsheet[1][1] == nil ? '' : paramsheet[1][1].value.to_s
h_QTR = paramsheet[2][1] == nil ? '' : paramsheet[2][1].value.to_s
h_BANK = paramsheet[3][1] == nil ? '' : paramsheet[3][1].value.to_s

i, j = 0, 0         # i refers to the index of rows, j refers to the index of columns

worksheet.each { |row|      # iterate all rows in the worksheet
    if i == 0
        i += 1
        next            # ignore the first row because the first row contains titles
    end

    # read values from ith row in worksheet : e.g worksheet[i][j].value is for value of the ith row & jth column of the sheet

    acct_number = worksheet[i][0] && worksheet[i][0].value

    branch = worksheet[i][2] && worksheet[i][2].value
    account = worksheet[i][3] && worksheet[i][3].value

    cash = worksheet[i][4] == nil ? 0.0 : worksheet[i][4].value.to_f
    credit_bal = worksheet[i][5] == nil ? 0.0 : worksheet[i][5].value.to_f
    commodities = worksheet[i][6] == nil ? 0.0 : worksheet[i][6].value.to_f
    equity = worksheet[i][7] == nil ? 0.0 : worksheet[i][7].value.to_f
    fixed_income = worksheet[i][8] == nil ? 0.0 : worksheet[i][8].value.to_f
    fx_cash = worksheet[i][9] == nil ? 0.0 : worksheet[i][9].value.to_f
    margin_loan = worksheet[i][10] == nil ? 0.0 : worksheet[i][10].value.to_f
    non_traditional = worksheet[i][11] == nil ? 0.0 : worksheet[i][11].value.to_f
    other = worksheet[i][12] == nil ? 0.0 : worksheet[i][12].value.to_f
    restricted_short_cash = worksheet[i][13] == nil ? 0.0 : worksheet[i][13].value.to_f

    #constants  : these constants will be defined on input excel file on version 2

    opn = 'OPN'
    column_G = 'YE 1Q19 Bank Balance'
    debits = 'D'
    credits = 'C'

    balmstr = 'BALMSTR'

    # for each row of input file, creating 6 rows for output

    #branch 
    for k in 0..6
        res_sheet.add_cell(j + k, 0, branch)    # res_sheet.add_cell ( x, y, value ) means that setting value to xth row and yth column of the sheet
    end

    #date
    for k in 0..6
        res_sheet.add_cell(j + k, 1, date)
    end

    #OOP
    for k in 0..6
        res_sheet.add_cell(j + k, 2, opn)
    end

    #account
    if account != nil
        res_sheet.add_cell(j, 3, account)
        res_sheet.add_cell(j + 1, 3, '123.' + account[3..-1])       # account[3..-1] : substring (from the 3rd character) of account
        res_sheet.add_cell(j + 2, 3, '125' + account[3..-1])

        suffix_index = account.reverse.index('-')               # get the last index of '-' in account string
        suffix_account = account[-suffix_index..-1]             # get the suffix of account string
        res_sheet.add_cell(j + 3, 3, '30000001-' + suffix_account)
        res_sheet.add_cell(j + 4, 3, '30000002-' + suffix_account)
        res_sheet.add_cell(j + 5, 3, '30000001-' + suffix_account)
    end

    #Amount

    res_sheet.add_cell(j, 4, cash + credit_bal + fx_cash)
    res_sheet.add_cell(j + 1, 4, non_traditional)
    res_sheet.add_cell(j + 2, 4, commodities + equity + fixed_income + margin_loan + other + restricted_short_cash)
    res_sheet.add_cell(j + 3, 4, 0)
    res_sheet.add_cell(j + 4, 4, 0)
    res_sheet.add_cell(j + 5, 4, cash + credit_bal + fx_cash + non_traditional + commodities + equity + fixed_income + margin_loan + other + restricted_short_cash)

    #Debits and Credits
    for k in 0..6
        if k < 3
            res_sheet.add_cell(j + k, 5, debits)
        else
            res_sheet.add_cell(j + k, 5, credits)
        end
    end

    #column_G
    for k in 0..6
        res_sheet.add_cell(j + k, 6, column_G)
    end

    #column_H
    for k in 0..6
        res_sheet.add_cell(j + k, 7, h_YR + h_QTR + h_BANK + (acct_number % 10000).to_s)        # acct_number % 10000 : get the last 4 digits, to_s : convert int to string
    end

    #column_I
    for k in 0..6
        res_sheet.add_cell(j + k, 8, k + 1)
    end

    #column_K
    for k in 0..6
        res_sheet.add_cell(j + k, 9, balmstr)
    end

    i += 1
    j += 6
}

dot_index = input_filename.reverse.index('.') + 1
output_filename = input_filename[0...-dot_index] + '-output.xlsx'


res_book.write(output_filename)   # save the excel file to output.xlsx