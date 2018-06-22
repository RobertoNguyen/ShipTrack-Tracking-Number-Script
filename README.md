# ShipTrack Tracking Number Script
Excel Script that retrieves the following return result from a tracking number:
- Status - returns current status of package
  - "Delivered"
  - "Out for Delivery"
  - "In Transit" - Last Updated Status & Location
- Delivery Date - date delivered, if available
- Schedule By - date package is scheduled to be delivered, if not already delivered
- Received By - name of person who signed for package, if available
- Shipped To - location package is being delivered to
- Shipping Service - type of ship service for package (3-Day, Ground, etc.)
- Label Created - date label for tracking number was created
- Origin - location the package came from
- TimeStamp - last time the program cache was updated

### **Excel Function Formula and Requirements**
- =ShipTrack(“tracking#”, “carrier”, “returnresult”, TRUE)
  - “tracking#” - must be quoted
  - "carrier" - available carriers: DHL, FedEx, UPS. Must be quoted
  - "returnresult" - list of available returns above must be quoted
  
### **Installation Method 1**
Drop TrackStatus.xlam in Microsoft Add-In folder:
- C:\Users\USERNAME\AppData\Roaming\Microsoft\AddIns
Activate the add-in in Excel. 
- File -> Options -> Add-ins -> Manage Excel Add-Ins -> Go -> Check box: Trackstatus -> Ok
Attach the add-in to ribbon bar for easy access
- File -> Options -> Quick Access Toolbar -> Choose commands from -> Macros -> Select the macro -> Add >> -> Ok

### **Installation Method 2**
Add button to worksheet and assign macro to it.
Activate “Developer” tab in Options
Insert Button (Form Control) and drag cell area to create custom sized button
Right click on button -> Assign Macro…
