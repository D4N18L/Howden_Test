import numpy as np
import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
import matplotlib.pyplot as plt


class DBConnection:
    """
    This class is used to connect to the database and execute statements
    """

    def __init__(self, connection_string):
        self.connection_string = connection_string
        self.engine = None

    def connect(self):
        """
        Create an engine instance with the connection string
        """
        try:
            self.engine = create_engine(self.connection_string)
            print("Connected to the database using the connection string: {}".format(self.connection_string))
        except Exception as e:
            print("Error connecting to the database: {}".format(e))

    def execute_statement(self, statement):
        """
        Execute a statement on the database using the engine instance
        """
        try:
            with self.engine.connect() as connection:
                connection.execute(statement)
                print("Executed the statement: {}".format(statement))
        except Exception as e:
            print("Error executing the statement: {}".format(e))

    def close(self):
        """
        Close the connection to the database
        """
        try:
            self.engine.dispose()
            print("Closed the connection to the database.")
        except Exception as e:
            print("Error closing the connection: {}".format(e))


class Problem1_A:
    """
    This class pulls statistical data from two tabs in the Excel file and combines them into a single DataFrame to be
    loaded into the database.
    """

    def __init__(self, file_path):
        self.file_path = file_path

    def read_data(self, sheet_name):
        """
        Read and combine data from an Excel sheet into a DataFrame.
        """

        years = pd.read_excel(self.file_path, sheet_name=sheet_name, usecols="A", nrows=12,
                              skiprows=4)  # Read the years from Column A, starting at row 6
        years = years.rename(columns={years.columns[0]: 'Year'})  # Rename the column to Year

        # Read the development months from C5:N5, which are the intervals after the first 12 months
        dev_months = pd.read_excel(self.file_path, sheet_name=sheet_name, usecols="C:N", nrows=1, skiprows=4).iloc[0]

        dev_months = [(i + 1) * 12 for i in
                      range(len(dev_months))]  # Calculate the development months by multiplying the intervals by 12

        # Read the loss incurred ratios from C6:N17
        loss_ratios = pd.read_excel(self.file_path, sheet_name=sheet_name, usecols="C:N", nrows=12, skiprows=4)

        # Prepare the DataFrame to combine all data
        combined_data = pd.DataFrame()

        for i, month in enumerate(dev_months):  # Loop through the development months

            temp_df = pd.DataFrame()
            temp_df['Year'] = years['Year'].values
            temp_df['DevelopmentMonth'] = month
            temp_df['LossIncurredRatio'] = loss_ratios.iloc[:, i].values
            temp_df['LineOfBusiness'] = sheet_name
            temp_df['Currency'] = 'EUR'
            temp_df['CompanyName'] = 'Howden Company'
            temp_df['DWCreatedDate'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            temp_df['DWCreatedBy'] = 'Daniel Tiboah-Addo'

            combined_data = pd.concat([combined_data, temp_df], ignore_index=True)

        # Reset index to clean up the DataFrame
        combined_data.reset_index(drop=True, inplace=True)

        return combined_data


class Problem1_B:
    """
    This class pulls booked and statistical data from two tabs , combine them into a single DataFrame and load them into
    the database.
    """

    def __init__(self, file_path):
        self.file_path = file_path

    def read_data(self, sheet_name):
        """
        Read and combine data from an Excel sheet into a DataFrame.
        """

        # Read the years and gross written premium from A6:B17
        years_and_gross_written_premium = pd.read_excel(self.file_path, sheet_name=sheet_name, usecols="A:B", nrows=12,
                                                        skiprows=4)
        years_and_gross_written_premium.columns = ['Year', 'GrossWrittenPremium']

        # Read the booked data from P6:S17
        booked_data = pd.read_excel(self.file_path, sheet_name=sheet_name, usecols="P:S", nrows=12, skiprows=4)
        booked_data.columns = ['EarnedPremium', 'PaidLosses', 'CaseReserves', 'IBNR']

        # Calculate the Ultimate Losses ratio by adding the Paid Losses + Case Reserves + IBNR
        booked_data['UltimateLossRatio'] = booked_data['PaidLosses'] + booked_data['CaseReserves'] + booked_data['IBNR']

        # Combine years and gross written premium with booked data
        combined_data = pd.concat([years_and_gross_written_premium, booked_data], axis=1)

        # Add the LineOfBusiness CompanyName, DWCreatedDate, DWCreatedBy columns to the DataFrame
        combined_data['LineOfBusiness'] = sheet_name
        combined_data['DWCreatedDate'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        combined_data['DWCreatedBy'] = 'Daniel Tiboah-Addo'

        return combined_data


class Problem1_C:
    """
    This class reads data from the booked table from any of the tabs and generate a graph.
    """

    def __init__(self, file_path):
        self.file_path = file_path

    def read_and_visualise(self, sheet_name):
        """
        Read and visualise data from an Excel sheet
        """

        # Read the years from A6:A17
        years = pd.read_excel(self.file_path, sheet_name=sheet_name, usecols="A", nrows=12, skiprows=4)
        years = years.rename(columns={years.columns[0]: 'Year'})

        # Read the booked data from P6:S17
        booked_data = pd.read_excel(self.file_path, sheet_name=sheet_name, usecols="P:S", nrows=12, skiprows=4)

        # Combine the years and booked data
        booked_data = pd.concat([years, booked_data], axis=1)
        booked_data.columns = ['Year', 'EarnedPremium', 'PaidLosses', 'CaseReserves', 'IBNR']

        # reset index to clean up the DataFrame
        booked_data = booked_data.reset_index()

        # Create a stacked bar chart
        ax = booked_data.plot(x='Year', kind='bar', stacked=True,
                              y=['PaidLosses', 'CaseReserves', 'IBNR'],
                              color=['blue', 'lightgreen', 'darkgreen'])

        # Add a line chart to the secondary y-axis that overlays the stacked bar chart with the earned premium
        booked_data['EarnedPremium'].plot(kind='line', color='lightblue', marker='*', linewidth=1, ax=ax,
                                          secondary_y=True, label='Earned Premium')

        # Customize ticks on the primary y-axis (left) to show percentages
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f'{y:.0%}'))

        # Customize ticks on the secondary y-axis (right) to increment by 200
        max_premium = booked_data['EarnedPremium'].max()
        ax.right_ax.set_yticks(np.arange(0, max_premium + 200, 200))
        ax.right_ax.set_ylim(0, max_premium + (200 - (max_premium % 200)))  # Adjust limit to the next multiple of 200

        ax.figure.legend(loc='lower center', bbox_to_anchor=(0.5, 0.01), ncol=4)

        plt.show()
