# gcp-slidegen

This program takes weekly update slides and collates them into one slide deck. It is an automation solution that allows users to quickly and easily create slide decks for client check-in calls using relevant updates since the last client meeting. It exists as a browser add-on that users can run from Google Sheets.

Benefits include:

For Management - centralized command, allows mass messaging, augment knowledge trasnfer, consistent execution of client experience, maximises utilisation, drives cross sell, lowers churn.

For Clients - Consultant-agnostic structured check-ins consistent throughout contract lifecycles, continued product updates, reference material as leave behind after check-in, awareness of marketing/training events, professional briefs attributable to the brand.

For Users - automation (5 min training, 1 click usage), knowledge preserved in centralized folders, new users onboard faster with standardized deck, content aleady created by other teams, less duplication of work - more focus on providing value.




Technical workflow:

This details the basic technical workflow, actions of the program, and the Google Drive folders touched.


Google Drive folders read:

USER’S top-level personal folder

‘Customers’ folder inside [top-level-folder] shared folder

‘Slide Generator’ folder inside [top-level-folder] shared folder

USER enters variables and then executes ‘Generate Slides’ method from menu bar of their SOURCE FILE

 

PROGRAM STARTS


Reads from Google Sheet file the variables 1. CUSTOMER NAME, 2. LAST MEETING DATE, 3. NEXT MEETING DATE, 4. DATA SOURCE FOLDERS ARRAY

 

CREATES NEW blank Google Slide file SLIDE DECK in USER’s top-level personal folder on Google Drive with naming convention ‘YYYY-MM-DD CUSTOMER NAME Check-In’

 

MOVES (via copying and then deleting original) blank SLIDE DECK from USER’S top-level personal folder on Google Drive to folder: [top-level-folder]/Customers/[CUSTOMER INITIAL]/CUSTOMER NAME/checkins

 

READS certain folders in [top-level-folder]/Slide Generator/slides and ADDS certain slides from Google Slide files there. Checks LAST MEETING DATE to limit which slides are added, based on date. Default slides: 1. Slide in folder ‘client’ with CUSTOMER NAME if one has been added (also, if the string ‘INSERTDATE’ is found on this slide, replaces this string with NEXT MEETING DATE), 2. In-date slides from ‘exec’ folder.

 

Then, based on contents of DATA SOURCE FOLDERS ARRAY, ADDS slides from other files in other folders e.g. if ‘news-sites’ is requested by the USER in SOURCE FILE, the PROGRAM will ADD slides from files in the ‘news-sites’ folder. Will also add a header slide called ‘news-sites’ from folder ‘header’ if it exists.

 

WRITES LAST GENERATED ON variable to USER’S SOURCE FILE


PROGRAM ENDS
