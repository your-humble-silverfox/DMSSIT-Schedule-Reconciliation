# Schedule Reconciliation

A script performing reconciliation of schedule and workload of Russian Technological University MIREA, verifying that student-groups have the right professors attached to them, as well as reporting any and all classes in the schedule without a professor assigned. Due to the nature of the project, majority of the comments and documentation for this is written in Russian language, however, an english-language readme is created and maintained for those interested in the project.

# Documentation
## Using the script

Being a python-script with a cli interface, it is launched from a terminal via the following command:

```bash
python schedule_reconciliation.py [path to schedule spreadsheet] [path to workload spreadsheet]
```

**Warning:** due to the nature of the openpyxl, schedule spreadsheet has to be a `.xlsx` file!!!