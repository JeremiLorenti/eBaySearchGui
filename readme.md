eBay Search Using Ruby
This is a simple GUI application built with Ruby that allows users to search for eBay listings and save the results to an Excel file. The application uses the Nokogiri gem to parse HTML from the eBay website and the Axlsx gem to create an Excel file.

Installation
To use this application, you will need to have Ruby installed on your computer. You can download Ruby from the official website.

You will also need to install the following gems:

Nokogiri
Open-URI
URI
Caxlsx
Win32ole
Tk
You can install these gems by running the following command in your terminal:

Copy codegem install nokogiri open-uri uri caxlsx win32ole tk
Usage
To use the application, simply run the 
search_window.rb
 file in your terminal:

Copy coderuby search_window.rb
This will open the GUI window where you can enter your search term, select the listing type (current or sold), and choose a file name to save the results to.

Once you have entered your search criteria, click the "Search" button to retrieve the eBay listings. The application will display the results in the GUI window and ask if you want to save the data to a file.

If you choose to save the data, the application will open a Windows Explorer window where you can choose the save location and file name. The data will be saved to an Excel file in the selected location.

Contributing
If you would like to contribute to this project, feel free to submit a pull request.
