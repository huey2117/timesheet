import os
import pandas as pd
import numpy as np
import xlsxwriter
from datetime import timedelta, datetime, date


def get_payperiod_dates(start_date):
    end_date = start_date + timedelta(days=13)
    return f'{start_date.strftime("%m.%d.%Y")}-{end_date.strftime("%m.%d.%Y")}'


def get_row_dates(start_date):
    return [(
        start_date + timedelta(days=i)).strftime('%a %b %d')
            for i in range(7)
            if not (start_date + timedelta(days=i)).weekday() == 6
    ]


def main():
    # Initial definitions
    start_date = datetime.strptime('12.13.2020', '%m.%d.%Y').date()
    first_df_row = 1
    second_df_row = 15

    # initialize writer
    writer = pd.ExcelWriter('HW Timesheet 2021.xlsx', engine='xlsxwriter')

    while start_date + timedelta(days=14) < date(2022, 1, 1):
        # Begin First Data Frame per page
        main_headers = ['Outpatient - EN',
                        'Outpatient - Other',
                        'Contract Services - School',
                        'Contract Services - Travel Time',
                        'Contract Services - Mileage',
                        'Time Off - PTO',
                        'Time Off - Holiday'
                        ]
        main_index = pd.date_range(
            start_date + timedelta(days=1),
            periods=6,
            freq='D').strftime('%a %b %d')
        main_df = pd.DataFrame(columns=main_headers, index=main_index)
        a = main_df.columns.str.split(' - ', expand=True).values
        main_df.columns = pd.MultiIndex.from_tuples([x for x in a])
        sheet = get_payperiod_dates(start_date)
        main_df.to_excel(writer, sheet, index=True, startrow=first_df_row)
        writer.sheets[sheet].set_row(3, None, None, {'hidden': True})

        # Begin 2nd data frame per page
        sec_index = pd.date_range(
            start_date + timedelta(days=8),
            periods=6,
            freq='D').strftime('%a %b %d')
        sec_df = pd.DataFrame(columns=main_headers, index=sec_index)
        b = sec_df.columns.str.split(' - ', expand=True).values
        sec_df.columns = pd.MultiIndex.from_tuples([x for x in b])
        sec_df.to_excel(writer, sheet, index=True, startrow=second_df_row)
        writer.sheets[sheet].set_row(17, None, None, {'hidden': True})

        # define some formatting
        background_color = '#FFFFFF'
        font_color = '#000000'
        bold = writer.book.add_format(
            {
                'bold': 1,
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            }
        )
        string_format = writer.book.add_format(
            {
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1,
                'align': 'center'
            }
        )

        number_format = writer.book.add_format(
            {
                'num_format': '0.00',
                'font_color': font_color,
                'bg_color': background_color,
                'align': 'center',
                'border': 1
            }
        )

        integer_format = writer.book.add_format(
            {
                'num_format': '0',
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1,
                'top': 2,
                'align': 'center'
            }
        )

        right_border_format = writer.book.add_format(
            {
                'left': 1
            }
        )

        # Sheet summary:
        # Outpatient Hours: =sum()
        # Contract Services Hours: =sum()
        # Time Off: =sum()
        # Total Hours: = sum()
        summary_cols = ['Summary', 'Hours']
        summary_rows = ['Outpatient Hours', 'Contract Services Hours', 'Time Off', 'Total Hours']
        summary_forms = ['=sum(B11,C11,B25,C25)',
                         '=sum(D11:F11,D25:F25)',
                         '=sum(G11,H11,G25,H25)',
                         '=sum(B11:H11,B25:H25)'
                         ]
        summary_df = pd.DataFrame(columns=summary_cols)
        for r, f in zip(summary_rows, summary_forms):
            summary_df = summary_df.append(
                pd.Series(
                    [r, f],
                    index=summary_cols
                ),
                ignore_index=True
            )

        summary_df.to_excel(writer, sheet, index=False, startrow=27)

        # Format the data
        writer.sheets[sheet].set_row(31, 15, bold)
        writer.sheets[sheet].set_column('A:A', 20, string_format)
        writer.sheets[sheet].set_column('B:H', 10, number_format)
        writer.sheets[sheet].set_column('I:I', 10, right_border_format)

        writer.sheets[sheet].write(0, 0, start_date.strftime('%m.%d.%Y'), bold)
        writer.sheets[sheet].write(14, 0, (start_date + timedelta(days=7)).strftime('%m.%d.%Y'), bold)
        # NEED TO UN-HARDCODE THIS 10
        writer.sheets[sheet].write(10, 0, 'Total', bold)
        writer.sheets[sheet].write(24, 0, 'Total', bold)
        for i, col in enumerate(['B', 'C', 'D', 'E', 'F', 'G', 'H']):
            writer.sheets[sheet].write(10, i + 1, f'=sum({col}4:{col}9)', number_format)
            writer.sheets[sheet].write(24, i + 1, f'=sum({col}19:{col}24)', number_format)

        start_date += timedelta(days=14)

    writer.save()


if __name__ == '__main__':
    main()
