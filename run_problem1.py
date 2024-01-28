import pandas as pd

from load_config import load_config
from Problem1 import DBConnection, Problem1_A, Problem1_B, Problem1_C


def main():
    # Load the configuration
    config = load_config()
    connection_string = config['db_url']

    # Establish a database connection
    db_conn = DBConnection(connection_string)
    db_conn.connect()

    # ------------------ Problem 1A ------------------
    pb1a = Problem1_A('Howden_CompanyXYZ_2021_Data.xlsx')
    data_gl_np = pb1a.read_data('GL-np')
    data_ma_np = pb1a.read_data('MA-np')

    # Combine the data
    combined_data = pd.concat([data_gl_np, data_ma_np], ignore_index=True)

    print("Combined data - Problem 1A:")
    print(combined_data)

    # Load the data to the database
    combined_data.to_sql(name='factstatistical', con=db_conn.engine, if_exists='replace', index=False)

    print("Finished loading factstatistical data to the database.")

    # ------------------ Problem 1B ------------------
    pb1b = Problem1_B('Howden_CompanyXYZ_2021_Data.xlsx')
    problem1b_gl_np = pb1b.read_data('GL-np')
    problem1b_ma_np = pb1b.read_data('MA-np')

    # Combine the data
    combined_data = pd.concat([problem1b_gl_np, problem1b_ma_np], ignore_index=True)
    print("Combined data - Problem 1B:")
    print(combined_data)

    # Load the data to the database
    problem1b_gl_np.to_sql(name='factdata', con=db_conn.engine, if_exists='replace', index=False)

    print("Finished loading factdata data to the database.")

    # ------------------ Problem 1C ------------------
    pb1c = Problem1_C('Howden_CompanyXYZ_2021_Data.xlsx')
    pb1c.read_and_visualise('GL-np')

    print("Finished visualising the data.")

    # Close the database connection
    db_conn.close()

    exit()


if __name__ == "__main__":
    main()
