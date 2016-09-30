# pyssagov
Convert Social Security Statement to Excel Salary Trend Chart

Visit the [Social Security Website](https://www.ssa.gov) and SignUp.  Export your earnings to the xml format and note the file's location.
### Installation
```
pip install -r requirements.txt
```

### Usage
```
python convert_to_excel.py --file="/path-to/YYYY_Your_Social_Security_Statement_Data.xml"

```

The file will output to this app directory on the local machine as an Excel (xlsx) file.
