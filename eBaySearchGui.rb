require 'nokogiri'
require 'open-uri'
require 'uri'
require 'caxlsx'
require 'win32ole'
require 'tk'

class SearchWindow
  attr_accessor :search_entry

  def initialize
    # Create the main window
    @root = TkRoot.new {title "eBay Search Using Ruby"}

    # Create the search term label and entry
    search_label = TkLabel.new(@root) {text "Search Term:"; grid('row'=>0, 'column'=>0, 'sticky'=>'w')}
    @search_entry = TkEntry.new(@root) {grid('row'=>0, 'column'=>1, 'sticky'=>'w')}

    # Create the listing type label and radio buttons
    listing_label = TkLabel.new(@root) {text "Listing Type:"; grid('row'=>1, 'column'=>0, 'sticky'=>'w')}
    @listing_type = TkVariable.new
    current_radio = TkRadioButton.new(@root) {text "Current Listings"; variable @listing_type; value 'c'; grid('row'=>1, 'column'=>1, 'sticky'=>'w')}
    sold_radio = TkRadioButton.new(@root) {text "Sold Listings"; variable @listing_type; value 's'; grid('row'=>1, 'column'=>2, 'sticky'=>'w')}

    # Create the save file label and entry
    save_label = TkLabel.new(@root) {text "Save File Name:"; grid('row'=>2, 'column'=>0, 'sticky'=>'w')}
    @save_entry = TkEntry.new(@root) {grid('row'=>2, 'column'=>1, 'sticky'=>'w')}

    # Create the search button
    search_button = TkButton.new(@root) {text "Search"; grid('row'=>3, 'column'=>0, 'columnspan'=>3)}

    # Create the status label
    @status_label = TkLabel.new(@root) {text "Ready"; grid('row'=>4, 'column'=>0, 'columnspan'=>3)}

    # Bind the search function to the search button
    search_button.command {search}

    # Start the Tk event loop
    Tk.mainloop
  end

  def search
    # Get the search term and listing type from the GUI
    search_term = @search_entry.get
    listing_type = @listing_type.value

    # Construct the eBay search results URL using the search term and listing type
    url = "https://www.ebay.com/sch/i.html?_nkw=#{URI::DEFAULT_PARSER.escape(search_term)}"
    if listing_type == 's'
      url += '&LH_Sold=1'
    end

    # Use Nokogiri to parse the HTML from the URL
    html = URI.open(url).read
    doc = Nokogiri::HTML(html)

    # Find all the listings on the page
    listings = doc.css('.s-item')

    # Initialize the data array
    data = [['Item Name', 'Condition', 'Sold Price', 'Shipping']]

    # Loop through each listing and extract the data we want
    listings.each do |listing|
      # Extract the title, price, and shipping cost
      title = listing.css('.s-item__title').text.strip
      condition = listing.css('.SECONDARY_INFO').text.strip
      price = listing.css('.s-item__price').text.gsub('US $', '$').gsub(' to ', '-')
      shipping = listing.css('.s-item__shipping.s-item__logisticsCost').text.strip

      # Split the shipping text into type and amount
      if shipping == 'Free shipping'
        shipping_type, shipping_amount = 'Free Shipping', nil
      elsif shipping.include?('to')
        shipping_type, shipping_amount = 'Range', shipping.gsub(' shipping','').strip
      else
        shipping_type, shipping_amount = 'Amount Charged', shipping.gsub('+','').strip
      end

      # Add the data to the array
      data << [title, condition, price, "#{shipping_type} - #{shipping_amount}"]
    end

    # Create the Excel file
    package = Axlsx::Package.new
    listing_type = @listing_type.value
    workbook = package.workbook
    worksheet = workbook.add_worksheet(:name => "eBay Listings")

    data.each do |row|
      worksheet.add_row row
    end

    # Ask the user if they want to save the file
    save_file = Tk.messageBox('type' => 'yesno', 'icon' => 'question', 'message' => 'Do you want to save the data to a file?')

    if save_file == 'yes'
      # Open Windows Explorer to let the user choose the save location and file name
      shell = WIN32OLE.new('Shell.Application')
      save_dialog = shell.BrowseForFolder(0, "Select the folder to save the file in", 0, 0)
      if save_dialog != nil
        save_folder = save_dialog.Items().Item().Path
        file_name = "#{@save_entry.get}.xlsx"
        file_path = File.join(save_folder, file_name)

        # Save the Excel file
        package.serialize(file_path)

        @status_label.text = "Data saved to #{file_path}"
      end
    else
      @status_label.text = "Data not saved"
    end
  end
end

# Create an instance of the SearchWindow class
SearchWindow.new