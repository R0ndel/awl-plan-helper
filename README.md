# How to use

1. Download the monthly schedule from: https://www.wojsko-polskie.pl/awl/dziekanat-rozklad-zajec/
2. Update the `INPUT_FILE` variable with the name of downloaded excel sheet
3. Choose your major / generate .ics files for all majors

- `get_calendar_schemes(wb)` - creates an .ics file for each respective year for majors listed in `MAJORS`
- `get_major_calendar_scheme(wb, "<major>", <year>)` - creates an .ics file for supplied major and year

4. Upload the file to your calendar app
