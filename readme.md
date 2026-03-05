# Schedule Reconciliation

A FastAPI app performing reconciliation of schedule and workload of Russian Technological University MIREA, verifying that student-groups have the right professors attached to them, as well as reporting any and all classes in the schedule without a professor assigned. Due to the nature of the project, majority of the comments and documentation for this is written in Russian language, however, an english-language readme is created and maintained for those interested in the project.

# Documentation
## Available API requests:

- `reconcile` - POST-Request for full schedule reconciliation, receiving two tables in excel formats for reconciliation.
- `reconcile_mismatch` - POST-Request, returning only mismatches
- `reconcile_no_prof` - POST-Request, returning a list of classes without an assigned professor
