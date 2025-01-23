# Event Planner App

Hi there! Welcome to my Event Planner App. This project is a proof-of-concept designed to help small businesses manage their clients, events, and vendors in a simple, efficient way. It combines an Access database with an Excel frontend and VBA automation to make event planning smooth.

---

## Why I Built This
I wanted to create the tool that provides a bridge between database management and an easy-to-use interface. The app was designed for small businesses like wedding planners, corporate event organizers, or even freelancers who juggle multiple projects and clients. 

It’s a practical example of how Access, Excel, and VBA can work together to solve real-world problems.

---

## Features

### Access Database
- The backbone of this app is the Access database, which includes:
  - **Clients Table**: Keeps track of client details like name and contact info and we also assign an ID to the clients.
  - **Events Table**: Stores information about events, including type, date, venue, budget and vendor details.
  - **Vendors Table**: Manages vendors, their specialties, and contact details.
- I’ve also included three useful queries to demonstrate how data can be analyzed:
  1. **Events by Date**: See all upcoming events sorted by date.
  2. **Top Vendors**: Find the vendors hired the most frequently.
  3. **Event Types**: Get a count of events group by type (e.g., weddings, corporate, otherss).

### Excel Frontend
- The Excel workbook is designed to make working with the database easier:
  - **Data Display Tab**: Imports event data from the database and uses conditional formatting to highlight upcoming events.
  - **Analysis Tab**: Created a pivot table to Provides insights  (e.g., total budgets by event type or vendor).
  - **Data Input Tab**: This is a user-friendly form where you can add new events directly into the database.

### VBA Middleware
- VBA ties everything together. Here’s what it does:
  - **Import Data**: Pulls event data from the Access database into the Excel sheet.
  - **Add Event**: We can able to add new events via the Data Input Tab and updates the database automatically.
  - Error handling is built in to ensure the app runs smoothly!

---

## How to Use the App
1. Download both files:
   - `EventPlanner.accdb`: The Access database.
   - `EventPlanner.xlsm`: The Excel file with VBA macros.
2. Open the Excel file and make sure macros are enabled.
3. Use the **Data Input Tab** to add new events. Just fill in the fields and hit the button!
4. Check out the **Data Display Tab** to see all events.
5. Head over to the **Analysis Tab** to explore summaries and insights.

---

## What You’ll Need
To use this app, you’ll need:
- **Microsoft Access** (2016 or later).
- **Microsoft Excel** (2016 or later) with macros enabled.

---

## What’s Next?
This app is just a starting point. In the future, it could be scaled to include features like:
- Multi-user access.
- Integration with online platforms for real-time data updates.
- More advanced reporting features.

---

## Let’s Connect
If you have questions, feedback, or ideas for improvement, feel free to reach out! You can contact me at virajpahade1008@gmail.com. I’d love to hear from you.

---

Thanks for checking out my project. I hope it inspires you to explore how these tools can be combined in creative ways!
