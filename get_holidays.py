import requests
from datetime import datetime, timedelta
import locale

# Set the locale to German
locale.setlocale(locale.LC_TIME, "de_DE.UTF-8")


# Function to get all holiday dates within the range and format them
def get_holiday_dates(year):
    # URLs for school holidays and public holidays
    school_holidays_url = f"https://openholidaysapi.org/SchoolHolidays?countryIsoCode=DE&languageIsoCode=DE&validFrom={year}-01-01&validTo={year}-12-31&subdivisionCode=DE-BE"
    public_holidays_url = f"https://openholidaysapi.org/PublicHolidays?countryIsoCode=DE&languageIsoCode=DE&validFrom={year}-01-01&validTo={year}-12-31&subdivisionCode=DE-BE"

    # Initialize empty lists to store holiday dates
    school_holidays_dates = []
    public_holidays_dates = []
    print("Try to get holidays data from the API")

    try:
        # Get school holiday data
        response_school_holidays = requests.get(school_holidays_url)
        response_school_holidays.raise_for_status()  # Raise an error for bad responses
        school_holidays_data = response_school_holidays.json()
        # Extract and format school holiday dates
        for holiday in school_holidays_data:
            start_date = holiday["startDate"]
            end_date = holiday["endDate"]
            # Convert start_date to a datetime object
            current_date = datetime.strptime(start_date, "%Y-%m-%d")
            # Format and append each day within the date range, excluding weekends
            while current_date <= datetime.strptime(end_date, "%Y-%m-%d"):
                if current_date.weekday() < 5:  # Check if it's a weekday (0-4 are Monday to Friday)
                    formatted_date = current_date.strftime("%d.%m")
                    school_holidays_dates.append(formatted_date)
                current_date += timedelta(days=1)
        print("School data retrieved from the API successfully.")
    except requests.exceptions.RequestException as e:
        print(f"Fail to fetch data from school holiday.\nCheck Internet!")

    try:
        # Get public holiday data
        response_public_holidays = requests.get(public_holidays_url)
        response_public_holidays.raise_for_status()  # Raise an error for bad responses
        public_holidays_data = response_public_holidays.json()
        # Extract and format public holiday dates
        for holiday in public_holidays_data:
            start_date = holiday["startDate"]
            # Convert start_date to a datetime object
            current_date = datetime.strptime(start_date, "%Y-%m-%d")
            # Format and append each public holiday date
            formatted_date = current_date.strftime("%d.%m")
            public_holidays_dates.append(formatted_date)
        print("Holiday data retrieved from the API successfully.")
    except requests.exceptions.RequestException as e:
        print(f"Fail to fetch data from public holiday.\nCheck Internet!")

    return school_holidays_dates, public_holidays_dates


# Example usage
if __name__ == "__main__":
    year = "2024"
    school_holidays, public_holidays = get_holiday_dates(year)
    if school_holidays:
        print("School Holidays (Excluding Weekends):")
        for date in school_holidays:
            print(date)
    if public_holidays:
        print("Public Holidays:")
        for date in public_holidays:
            print(date)
