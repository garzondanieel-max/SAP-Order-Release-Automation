# SAP Order Release Automation

## Overview

This project is a Python-based desktop application designed to automate the mass release of SAP orders using SAP GUI Scripting. It reduces manual effort, minimizes human error, and significantly improves operational efficiency in repetitive SAP workflows.

The tool reads order data from an Excel file, connects to SAP, filters relevant records, and executes automated release actions through a user-friendly graphical interface.

---

## Business Problem

In traditional SAP operations, releasing multiple orders manually is:

* Time-consuming
* Repetitive
* Inefficient for high-volume workloads

This process required operators to navigate SAP screens repeatedly for each order, impacting productivity and operational speed.

---

## Solution

This automation streamlines the process by:

* Reading order data from an Excel file
* Automating SAP login via SAP GUI Scripting
* Navigating SAP interface programmatically
* Filtering and selecting relevant orders
* Executing bulk release actions
* Handling SAP popups and session interruptions
* Providing a simple GUI for end users

---

## Technologies Used

* Python
* SAP GUI Scripting
* Pandas
* Tkinter (GUI)
* Pyperclip
* Win32 COM (SAP integration)
* Excel (input data source)

---

## Key Features

* Graphical user interface for easy operation
* Excel-based input processing
* Automated SAP session handling
* Bulk order processing
* Popup and error handling
* Clipboard integration for SAP filtering
* Time-efficient execution of repetitive tasks

## Input Example

The system uses an Excel file with the following structure:

| Pedido |
| ------ |
| 10001  |
| 10002  |
| 10003  |

---

## Impact

* Reduced manual effort in SAP order processing
* Improved execution speed for bulk operations
* Increased operational accuracy
* Enhanced user productivity through automation

---

## Notes

* This project is intended for educational and portfolio purposes.
* Sample data included is anonymized or fictional.
* The automation logic can be adapted to similar SAP workflows.

---

## Author

Developed by Daniel Garzón
Focused on process automation, data analysis, and enterprise efficiency solutions.



