# Sullivan-Automation-Backend
This is a series of Javascript automation tools made for google sheets that I made. It automates the inventory tracking and integrates with a dashboard service, resuling lab assistant overhead and empowering students to 3D print more!

## Queue Automation
The basic idea is this:
1. User submits a form, this populates the things in a row that they entered.
2. onFormSubmit sets the status to "In Queue" and adds VLOOKUP formulas.
3. When the status changes to "In Progress" or "Flagged/Completed", a timestamp is added.
4. When marked as "Completed" or "Flagged," the script acquires a lock, sends the email, archives the data, and immediately deletes the original row.

