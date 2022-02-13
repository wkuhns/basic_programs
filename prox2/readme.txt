The PROX.TXT file has exactly ten sections, one for each menu. Each section starts with a number indicating the number of display pages for the following section, followed by two quoted lines for each display page.
 
Each section may have up to 20 pages.

Only the contents and size of the menus may be changed. The logic that controls navigation is hard-coded.

The sections are arranged as follows:

Toplevel Menu
System Outputs
Present Status
Present Faults
Fault History top menu
Fault History, most recent leg (leg 75)
Fault History, most recent -1 leg (leg 74)
Fault History, most recent -2 leg (leg 73)
Fault History, most recent -3 leg (leg 72)
Fault History, most recent -4 leg (leg 42)

The first two display pages in the Fault History top menu do not branch to anything. The third page branches to the 'most recent leg' menu, the fourth to the 'most recent -1', and so on. The leg numbers are not part of the control logic - they are just display text.
