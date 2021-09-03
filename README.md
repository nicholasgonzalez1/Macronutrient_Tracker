# Macronutrient Tracker
<!-- TABLE OF CONTENTS -->
<details open="open">
  <summary>Table of Contents</summary>
  <ol>
    <li><a href="#about-the-project">About the Project</a></li>
    <li><a href="#built-with">Built With</a></li>
    <li><a href="#getting-started">Getting Started</a></li>
    <li><a href="#main-tab">Main Tab</a></li>
    <li><a href="#menu-tab">Menu Tab</a></li>
<!-- TO BE USED LATER
    <li>
      <a href="#user-interface-walkthrough">User Interface Walkthrough</a>
      <ul>
        <li><a href="#search-criteria-and-filters">Search Criteria and Filters</a></li>
        <li><a href="#view-more-flight-information">View More Flight Information</a></li>
        <li><a href="#reschedule-a-flight">Reschedule a Flight</a></li>
        <li><a href="#cancel-a-flight">Cancel a Flight</a></li>
      </ul>
    </li>
    <li><a href="#acknowledgements">Acknowledgements</a></li>
-->
  </ol>
</details>

<!-- ABOUT THE PROJECT -->
## About the Project

This spreadsheet is a personal project of mine which I gained inspiration from trying to transform my eating habits. After searching the internet for advice on how to improve my overall diet, I found that a key component of losing weight (and in turn, losing body fat - one of my primary goals) was eating in a calorie deficit. This means consuming less calories than you expend throughout the day. However, if I was going to truly implement this eating style, I needed to figure out how to stay full throughout the day and not get hungry. The solution to this problem was eating high-volume, low-calorie meals. 
<br><br>
I turned to scouring the internet once again for recipes that helped me abide by this rule. Between measuring exact serving sizes needed to make several recipes and keeping track of all my meals/snacks throughout the day, I felt that I needed a organizational tool to manage all details. Other applications such as MyFitnessPal offered convenient solutions, but I wanted to construct a tool that catered to my own styles and preferences. It was then that I decided to build this spreadsheet! I concluded that by developing the application myself, I'd be able to merge my fitness endeavors with my data science interests.
<br><br>
The spreadsheet is still a work in progress, but so far offers the ability to search branded and common foods from an online nutritional database, store the queried results into the spreadsheet, and organize meals by macronutrients and daily caloric intake.
<br><br>
<kbd>
<img src="https://github.com/nicholasgonzalez1/Macronutrient_Tracker/blob/main/images/user_interface.png?raw=true" width="700">
</kbd><br>

<!-- BUILT WITH -->
## Built With
The following the softwares and languages were used to implement this spreadsheet.
* Excel
* VBA
* Python

Excel was chosen to build the user interface since it was the platform I felt most comfortable using at the time, in addition to the simplicity of distributing it to other users as well. VBA was used to implement most of the background functionality of the spreadsheet, while Python scripts were developed in order to utilize API's which retrieved requested nutrition facts from an online database.

<!-- GETTING STARTED -->
## Getting Started
The current version of the user interface can be downloaded off [here](https://github.com/nicholasgonzalez1/Macronutrient_Tracker/blob/main/MACROS.xlsm). This Excel file **must** be downloaded as a *macro-enabled workbook*. The [menuScripts.py file](https://github.com/nicholasgonzalez1/Macronutrient_Tracker/blob/main/menuScripts.py) must be also downloaded and stored in the same folder location as the MACROS.xlsm file. 
<br><br>
Once the Excel file is downloaded, the xlWings add-on will need to be downloaded and added to the Excel application. xlWings is an open source package that serves as the bridge between Excel and Python. It comes pre-installed on Anaconda, but still needs to be enabled within the Excel application. Currently, xlWings only operates on Windows and macOS. The xlWings documentation offers further instructions on its [installation](https://docs.xlwings.org/en/stable/installation.html) and [add-in settings](https://docs.xlwings.org/en/stable/addin.html#xlwings-addin).
<br><br>
Certain features of the spreadsheet require an API key, which can be easily obtained for free on the [Nutritionix website](https://www.nutritionix.com/business/api). Registering for the free API key gives users access to 50 API calls per day. More information is provided [below]() on inputting the key into the spreadsheet and using the API.

## Main Tab

## Menu Tab
