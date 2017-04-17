class PqwExcel < ApplicationRecord
  def self.update_analysis
    #Create a new package and workbook
    package = Axlsx::Package.new
    workbook = package.workbook
    
    ##############################################################
    ##############################################################
    ############ALTER THIS PART BEFORE RUNNING####################
    ##############################################################
    ##############################################################
    #################
    #STEPS BEFORE RUNNING
    #1. Change start and end month/year to desired dates
    #2. Check SaS files noted for any change in pqw lists
    #3. Run program
    #################
      
      @start_month = 7
      @start_year = 2016
      @end_month = 2
      @end_year = 2017
    
    #################
    
    ##CHECK SAS FILE PQW_REPORT_MAILERS for additions/subtractions to pqw list
    ##For easier copy/paste, in excel use formula =CHAR(34)&cell&CHAR(34)&","
    @geps_pqws = [
      "All",
      "ACCESS WORLDWIDE PQW",
      "APC POSTAL LOGISTICS, LLC",
      "BROKER'S WORLDWIDE",
      "DHL ECOMMERCE",
      "GLOBEGISTICS PQW",
      "INTERNATIONAL BONDED COURIERS PQW",
      "MAIL SERVICES, INC.",
      "NIPPON EXPRESS USA, INC.",
      "PITNEY BOWES INTL",
      "POST-EDGE INTERNATIONAL",
      "RR DONNELLEY LOGISTICS",
      "UPS MAIL INNOVATIONS"
      ]
    
    ##CHECK SAS FILE P1_REPORT_MAILERS for addtions/subtractions to pqw list
    ##For easier copy/paste, in excel use formula =CHAR(34)&cell&CHAR(34)&","
    @pqws = [
      "All",
      "360 DISTRIBUTION",
      "ACCESS WORLDWIDE PQW",
      "AMERICAN INTERNATIONAL MAILING, INC.",
      "APC POSTAL LOGISTICS, LLC",
      "ARROWMAIL PRESORT COMPANY INC",
      "BROKER'S WORLDWIDE",
      "CHAMPION WORLDWIDE SOLUTIONS LLC",
      "DHL ECOMMERCE",
      "DST MAILING SERVICES INC.",
      "GEOPOST CORPORATIONS",
      "GLOBAL POSTAL SOLUTIONS",
      "GLOBEGISTICS PQW",
      "INTERNATIONAL BONDED COURIERS PQW",
      "INTERNATIONAL DELIVERY SOLUTIONS",
      "INTERNATIONAL TRANSPORT ACQUISITION, INC.",
      "MAIL ON THE MOVE",
      "MAIL SERVICES, INC.",
      "NIPPON EXPRESS USA, INC.",
      "ONTRAC INTERNATIONAL",
      "PITNEY BOWES INTL",
      "PITNEY BOWES PRESORT SERVICE, INC.",
      "POST-EDGE INTERNATIONAL",
      "RR DONNELLEY LOGISTICS",
      "SKYPOSTAL",
      "THREE DOG LOGISTICS",
      "UNIKTRANS CORPORATION",
      "UPS MAIL INNOVATIONS",
      "WORLDNET-SHIPPING USA, INC.",
      ]
    
    @products = [
      "All",
      "IPA Non-M-Bags-Mixed",
      "ISAL Non-M-Bags-Mixed",
      "IPA M-Bags",
      "IPA Non-M-Bags-Flats",
      "IPA Non-M-Bags-Letters",
      "IPA Non-M-Bags-Packets",
      "ISAL M-Bags",
      "ISAL Non-M-Bags-Flats",
      "ISAL Non-M-Bags-Letters",
      "ISAL Non-M-Bags-Packets",
      "ePacket",
      "PMEI",
      "PMI",
      "FCPIS"
      ]
    
    @p1_rates = [*1..19]
    @postal_one_rategroups = [
      "All",
      "Worldwide",
      @p1_rates
      ].flatten
    
    @geps_rates = [*2..17]
    @geps_rategroups = [
      "All",
      1,
      1.1,
      1.2,
      1.3,
      1.4,
      1.5,
      1.6,
      1.7,
      1.8,
      @geps_rates
      ].flatten
    
    @weights = [*1..70]
    @weightsteps = [
      "All",
      0.5,
      @weights
      ].flatten

    ##############################################################
    ##############################################################
    ##############################################################

    #Pull out months that will become columns for tables
    @start_date = Date.new(@start_year, @start_month, 1)
    @fiscal_end = Date.new(@start_year, 9, 15)
    @end_date = Date.new(@end_year, @end_month, 1)
    @date_arr = [@start_date.month.to_s]
    
    @current_year_count = 1
    @next_year_count = 0
    
    while @start_date < @end_date do
      @start_date += 1.months
      
      if @start_date < @fiscal_end
        @current_year_count += 1
      else
        @end_year = @start_year + 1
        @next_year_count += 1
      end
      
      @date_arr << @start_date.month.to_s
    end
    
    @start_yr_string = "FY" + @start_year.to_s[2..3]
    @end_yr_string = "FY" + @end_year.to_s[2..3]
    @current_arr = [@start_yr_string]
    (@current_year_count-1).times do
      
        @current_arr << " "  
    end
    
    if @next_year_count > 1 
      @next_arr = [@end_yr_string]
      (@next_year_count-1).times do
        @next_arr << " "  
      end
    elsif @next_year_count == 1
      @next_arr = [@end_yr_string]
    end
    
    @month_name = Date::MONTHNAMES[@end_month]
    @start_month_name = Date::MONTHNAMES[@start_month]
    @start_year_name = "FY" + @start_year.to_s[-2..-1]
    if @end_year != @start_year
      @end_year_name = "FY" + @end_year.to_s[-2..-1]
      @title_year = @start_year_name + "-" + @end_year_name
    else
      @title_year = @start_year_name
    end
    
    #Figuring out time in betweeen dates
    @month_difference = (@end_year * 12 + @end_month) - (@start_year * 12 + @start_month)
    
    ####STYLING FOR ALL TABS######
    workbook.styles do |s|
      @header = s.add_style b: true, sz: 14
      @bold_header = s.add_style sz: 8, b: true, :font_name => 'Arial'
      
      @tab_header = s.add_style b:true, sz: 16
      @tab_subheader = s.add_style b:true, sz:12
      
      @toc_row_header = s.add_style sz: 8, b: true, :font_name => 'Arial', :bg_color => '000000', :fg_color => 'ffffff'
      @toc_row_header_middle = s.add_style sz: 8, b: true, :font_name => 'Arial', :bg_color => '000000', :fg_color => 'ffffff', :border => { :style => :thin, :color =>"ffffff", :edges => [:right, :left] }
      @toc_row_styling = s.add_style sz: 8, alignment: {horizontal: :left}, :font_name => 'Arial'
      @toc_row_styling_middle = s.add_style sz: 8, :fg_color => '211aa5', alignment: {horizontal: :left}, :font_name => 'Arial', :border => { :style => :thin, :color =>"000000", :edges => [:right, :left] }
      
      @row_header = s.add_style sz: 10, b: true, :bg_color => '022256', :fg_color => 'ffffff', :font_name => 'Arial', :border => { :style => :thin, :color => 'ffffff', :edges => [:bottom] }
      @row_header_merge = s.add_style sz: 10, b: true, :bg_color => '022256', :fg_color => 'ffffff', :font_name => 'Arial', alignment: {horizontal: :center}, :border => { :style => :thin, :color => 'ffffff', :edges => [:bottom, :right] }
      @row_header_middle = s.add_style sz: 10, b: true, :bg_color => '022256', :fg_color => 'ffffff', :font_name => 'Arial', :border => { :style => :thin, :color => 'ffffff',  :edges => [:right, :left] }
      @row_header_right = s.add_style sz: 10, b: true, :bg_color => '022256', :fg_color => 'ffffff', :font_name => 'Arial', :border => { :style => :thin, :color => 'ffffff',  :edges => [:right, :bottom] }
      
      @tab_color1 = s.add_style :bg_color  => "faff00"
      @tab_color2 = s.add_style :bg_color  => "5cdef2"
      @tab_color3 = s.add_style :bg_color  => "b7f7b9"
      @tab_blue = s.add_style :bg_color => "3060ad"
      
      @plain_row = s.add_style sz: 8, alignment: {wrap_text: true, vertical: :top}
      @row_num = s.add_style sz: 6, i:true
      @row_styling = s.add_style sz: 8, alignment: {horizontal: :left}, :font_name => 'Arial', :border => { :style => :thin, :color =>"000000" }
      @row_styling_ital = s.add_style sz: 8, i: true, alignment: {horizontal: :left}, :font_name => 'Arial'
      @row_styling_middle = s.add_style sz: 8, alignment: {horizontal: :left}, :font_name => 'Arial', :border => { :style => :thin, :color =>"000000", :edges => [:right, :left] }
      @row_styling_rev = s.add_style sz: 8, :num_fmt => 7, alignment: {horizontal: :left}, :font_name => 'Arial', :border => { :style => :thin, :color =>"000000", :edges => [:right, :left] }
      @row_styling_vol = s.add_style sz: 8, :num_fmt => 3, alignment: {horizontal: :left}, :font_name => 'Arial', :border => { :style => :thin, :color =>"000000", :edges => [:right, :left] }
      @row_styling_per = s.add_style sz: 8, :num_fmt => 9, alignment: {horizontal: :left}, :font_name => 'Arial', :border => { :style => :thin, :color =>"000000", :edges => [:right, :left] }
      
      @row_styling_rev_total = s.add_style sz: 8, b:true, :num_fmt => 7, alignment: {horizontal: :left}, :font_name => 'Arial', :border => { :style => :thin, :color =>"000000", :edges => [:right, :left] }
      @row_styling_vol_total = s.add_style sz: 8, b:true, :num_fmt => 3, alignment: {horizontal: :left}, :font_name => 'Arial', :border => { :style => :thin, :color =>"000000", :edges => [:right, :left] }
      @row_styling_per_total = s.add_style sz: 8, b:true, :num_fmt => 9, alignment: {horizontal: :left}, :font_name => 'Arial', :border => { :style => :thin, :color =>"000000", :edges => [:right, :left] }
      @row_styling_total_text = s.add_style sz: 8, b:true, alignment: {horizontal: :left}, :font_name => 'Arial', :border => { :style => :thin, :color =>"000000", :edges => [:right, :left] }
      
      @last_row_middle = s.add_style sz: 8, alignment: {horizontal: :left}, :font_name => 'Arial', :border => { :style => :thin, :color =>"000000", :edges => [:right, :left, :bottom] }
      @last_row_rev = s.add_style sz: 8, :num_fmt => 7, alignment: {horizontal: :left}, :font_name => 'Arial', :border => { :style => :thin, :color =>"000000", :edges => [:right, :left, :bottom] }
      @last_row_vol = s.add_style sz: 8, :num_fmt => 3, alignment: {horizontal: :left}, :font_name => 'Arial', :border => { :style => :thin, :color =>"000000", :edges => [:right, :left, :bottom] }
      @last_row_per = s.add_style sz: 8, :num_fmt => 9, alignment: {horizontal: :left}, :font_name => 'Arial', :border => { :style => :thin, :color =>"000000", :edges => [:right, :left, :bottom] }
      
      @invisible = s.add_style :fg_color => 'ffffff'
      @wrap_text = s.add_style sz: 8, alignment: {wrap_text: true}
      @select = s.add_style sz: 8
      
    end
    
    #STYLE ARRAYS
      @style_merge = []
      @style_nil = []
      (@date_arr.length).times do
        @temp = @row_header_merge
        @style_merge << @temp
        @style_nil << nil
      end
      
      @date_style = []
      (@date_arr.length).times do
        @temp = @row_header_middle
        @date_style << @temp
      end
      
    #ROW STYLE ARRAYS
    @rev_style = []
    @rev_style_total = []
      (@date_arr.length).times do
        @temp = @row_styling_rev
        @rev_style << @temp
        @rev_style_total << @row_styling_rev_total
      end
      
    @vol_style = []
    @vol_style_total = []
      (@date_arr.length).times do
        @temp = @row_styling_vol
        @vol_style << @temp
        @vol_style_total << @row_styling_vol_total
    end
      
    @per_style = [@row_styling_per, @row_styling_per]
    
    @last_rev = []
      (@date_arr.length).times do
        @temp = @last_row_rev
        @last_rev << @temp
      end
      
    @last_vol = []
      (@date_arr.length).times do
        @temp = @last_row_vol
        @last_vol << @temp
    end
    
    @last_per = [@last_row_per, @last_row_per]
      
    ####TAB COLORS
    @tab1 = "faff00"
    @tab2 = "5cdef2"
    @tab3 = "b7f7b9"
    @tab_blu = "3060ad"
      
    ###VOLUME AND REVENUE HEADER SPACING
      @vol_arr = Array.new(@date_arr.length-1)
      @empty_arr  = Array.new(@date_arr.length)
      @rev_arr = Array.new(@date_arr.length-1)
      @vol_arr = @vol_arr.unshift("Calendar Month Volume")
      @rev_arr = @rev_arr.unshift("Calendar Month Revenue")
      
    #conditional format styling
      unprofitable = workbook.styles.add_style( :fg_color => "B22727", :type => :dxf, sz: 8 )
      profitable = workbook.styles.add_style( :fg_color => "076801" , :type => :dxf, sz: 8)
      
    #Conditoinal format arr
    @format_arr = ('C'..'Z').to_a
      if (@date_arr.length*2) > @format_arr.length
        @format_cont = ('A'..'Z').to_a
        @format_cont.map! { |word| "A#{word}" }
        @format_arr << @format_cont
        @format_arr.flatten!
      end
    
      @last_col = @format_arr[(@date_arr.length*2)+1]
      @second_last_col = @format_arr[(@date_arr.length*2)]
      
    ###TABLE OF CONTENTS PAGE###
    workbook.add_worksheet(name: "Table of Contents") do |sheet|
      
      ###Hide gridlines on page
      sheet.sheet_view.show_grid_lines= false
      
      ###Set up table of contents with hyperlinks and styling
      sheet.add_row []
      sheet.add_row [" ", "#{@month_name} " + "#{@end_year}" + " PQW Report - Table of Contents"], style: @header
      sheet.add_row []
      sheet.add_row []
      sheet.add_row ["", "Tab", "Title", "Description"], :style => [nil, @toc_row_header, @toc_row_header_middle, @toc_row_header]
      sheet.add_row ["", "1", "Table of Contents", "Summarizes contents of analysis"], :style => [nil, @toc_row_styling, @row_styling_middle, @toc_row_styling]
      sheet.add_row ["", "2", "Active Mailers", "Summary of active and inactive PQW mailers month over month"], :style => [nil, @toc_row_styling, @toc_row_styling_middle, @toc_row_styling]
      sheet.add_row ["", "3", "Products - ALL", "Summary of total volume and revenue data for products in PostalOne"], :style => [nil, @toc_row_styling, @toc_row_styling_middle, @toc_row_styling]
      sheet.add_row ["", "4", "Customers - ALL", "Summary of total IPA/ISAL and ePacket data for mailers in PostalOne"], :style => [nil, @toc_row_styling, @toc_row_styling_middle, @toc_row_styling]
      sheet.add_row ["", "5", "Postal One Rate Group Breakdown", "Summary of Postal One volume and revenue by rate group"], :style => [nil, @toc_row_styling, @toc_row_styling_middle, @toc_row_styling]
      sheet.add_row ["", "6", "GEPS Volume & Declines", "View top monthly GEPS declines and dynamic views of historical GEPS volume"], :style => [nil, @toc_row_styling, @toc_row_styling_middle, @toc_row_styling]
      sheet.add_row ["", "7", "GEPS Charts", "Overview of GEPS weightstep and rategroup volumes by PQW and product, filters other views"], :style => [nil, @toc_row_styling, @toc_row_styling_middle, @toc_row_styling]
      sheet.add_row ["", "8", "Monthly Volume by Weight Step", "Total Monthly GEPS Volume by Weight Step"], :style => [nil, @toc_row_styling, @toc_row_styling_middle, @toc_row_styling]
      sheet.add_row ["", "9", "Monthly Volume by Rate Group", "Total Monthly GEPS Volume by Rate Group"], :style => [nil, @toc_row_styling, @toc_row_styling_middle, @toc_row_styling]
      sheet.add_row ["", "10", "Monthly Volume by WS & RG", "GEPS Volume breakout by Weight Step and rate group"], :style => [nil, @toc_row_styling, @toc_row_styling_middle, @toc_row_styling]
      sheet.add_row ["", "11", "Monthly Volume Change WS & RG", "Change in volume broken out by Weight Step and rate group"], :style => [nil, @toc_row_styling, @toc_row_styling_middle, @toc_row_styling]
      sheet.add_row ["", "12", "Top Destinations by Monthly Vol", "Top destinations by monthly volume & chart"], :style => [nil, @toc_row_styling, @toc_row_styling_middle, @toc_row_styling]
      sheet.add_row ["", "13", "Top Customers by Monthly Vol", "Top PQW customers by monthly volume & chart"], :style => [nil, @toc_row_styling, @toc_row_styling_middle, @toc_row_styling]
      sheet.add_row ["", "14", "Rate Group Country Reference", "Rate groupings for countries by product"], :style => [nil, @toc_row_styling, @toc_row_styling_middle, @toc_row_styling]
      sheet.add_row ["", "15+", "Data Tables", "Data inputs to the analysis"], :style => [nil, @toc_row_styling, @toc_row_styling_middle, @toc_row_styling]
      sheet.add_row []
      sheet.add_row []
      sheet.add_row []
      sheet.add_row [" ","Tab Color Legend:"], style: @bold_header
      sheet.add_row [" ", "", "All Products Included"], :style => [nil, @tab_color1, @row_styling_ital]
      sheet.add_row [" ", "", "IPA/ISAL and CEP Only"], :style => [nil, @tab_color2, @row_styling_ital]
      sheet.add_row [" ", "", "GEPS/NMATS Products (PMI, PMEI, FCPIS) Only"], :style => [nil, @tab_color3, @row_styling_ital]
      
      sheet.add_hyperlink :location => "'Active Mailers'!A1", :ref => 'C7', :target => :sheet
      sheet.add_hyperlink :location => "'Products - ALL'!A1", :ref => 'C8', :target => :sheet
      sheet.add_hyperlink :location => "'Customers - ALL'!A1", :ref => 'C9', :target => :sheet
      sheet.add_hyperlink :location => "'Postal One Rate Group Breakdown'!A1", :ref => 'C10', :target => :sheet
      sheet.add_hyperlink :location => "'GEPS Volume & Declines'!A1", :ref => 'C11', :target => :sheet
      sheet.add_hyperlink :location => "'GEPS Charts'!A1", :ref => 'C12', :target => :sheet
      sheet.add_hyperlink :location => "'Monthly Volume by Weight Step'!A1", :ref => 'C13', :target => :sheet
      sheet.add_hyperlink :location => "'Monthly Volume by Rate Group'!A1", :ref => 'C14', :target => :sheet
      sheet.add_hyperlink :location => "'Monthly Volume by WS & RG'!A1", :ref => 'C15', :target => :sheet
      sheet.add_hyperlink :location => "'Monthly Volume Change WS & RG'!A1", :ref => 'C16', :target => :sheet
      sheet.add_hyperlink :location => "'Top Destinations by Monthly Vol'!A1", :ref => 'C17', :target => :sheet
      sheet.add_hyperlink :location => "'Top Customers by Monthly Vol'!A1", :ref => 'C18', :target => :sheet
      sheet.add_hyperlink :location => "'Rate Group Country Reference'!A1", :ref => 'C19', :target => :sheet
      sheet.add_hyperlink :location => "'PQW Report Data'!A1", :ref => 'C20', :target => :sheet
      
      
      sheet.column_widths 2, 14
    end
    
    ####ACTIVE MAILERS TAB####
    workbook.add_worksheet(name: "Active Mailers") do |mailer|
      ###Hide gridlines on page
      mailer.sheet_view.show_grid_lines= false
      mailer.sheet_pr.tab_color = @tab1
      
      mailer.add_row []
      mailer.add_row [" ", "Active Mailer Tracker"], :style => @tab_header
      mailer.add_row [" ", "Source: GEPS/NMATS & PostalOne, " + @title_year + " " + @start_month_name + "-" + @month_name], :style => @tab_subheader
      mailer.add_row []
      mailer.add_row []
      
      
      @letter_arr = ('C'..'Z').to_a
      @style_count_curr = []
      @style_count_next = []
      (@current_arr.length-1).times do
          @temp = @row_header_merge
          @style_count_curr << @temp
        end
        @style_count_curr << @row_header_right
        
      if @next_arr.empty?        
        mailer.add_row [" ", " ", @current_arr], :style => [nil, nil, @style_count_curr].flatten
      else
        (@next_arr.length).times do
          @temp = @row_header_merge
          @style_count_next << @temp
        end
        mailer.add_row ["", "", @current_arr, @next_arr].flatten, :style => [nil, nil, @style_count_curr, @style_count_next].flatten
        
        mailer.merge_cells "#{@letter_arr[@current_arr.length]}6:#{@letter_arr[@current_arr.length+@next_arr.length-1]}6"
        
      end
      
      
      
      mailer.merge_cells "C6:#{@letter_arr[@current_arr.length-1]}6"
      
      @style_date = []
      (@date_arr.length).times do
        @temp = @row_header_middle
        @style_date << @temp
      end
      mailer.add_row [" ", "PQW Mailer", @date_arr].flatten, :style => [nil, @row_header, @style_date].flatten
      
      
      @pqws.each_with_index do |pqw, index|
        if pqw != 'All'
         @letter_arr =('C'..'Z').to_a
         @static_arr =('C'..'Z').to_a
         @length_arr = [index, pqw]
         @style_arr = []
          (@date_arr.length).times do
            
            @letter = @letter_arr.shift
            @temp = %Q(=IF(IFERROR(VLOOKUP($B#{index+7},'Top Customers by Monthly Vol'!$B$7:$#{@static_arr[@date_arr.length]}$289,MATCH('Active Mailers'!#{@letter}$7,'Top Customers by Monthly Vol'!$B$7:$#{@static_arr[@date_arr.length]}$7,0),0),0)+IFERROR(VLOOKUP(UPPER($B#{index+7}), 'Postal One PQW Report'!$B$3:$#{@letter_arr[@date_arr.length]}$9999, MATCH("Totalvolume"&'Active Mailers'!#{@letter}$7, 'Postal One PQW Report'!$B$3:$#{@static_arr[@date_arr.length]}$3, 0), 0), 0)>0, "Yes", "No"))
            @length_arr << @temp
            @style_arr << @row_styling
          end
          mailer.add_row [@length_arr].flatten, :style => [@row_num, @row_styling, @style_arr].flatten
        end
      end
      
      mailer.column_widths 2, 35


      
      
      #conditional format styling
      non_active = workbook.styles.add_style( :fg_color => "B22727", :bg_color => "FFD1D1", sz: 8, :type => :dxf )
      active = workbook.styles.add_style( :fg_color => "076801" , :bg_color => "62ed5a", sz: 8, :type => :dxf)
    
      # Apply conditional formatting in the worksheet
      mailer.add_conditional_formatting("C:#{@letter}", { :type => :cellIs,
                                          :operator => :equal,
                                          :formula => "\"No\"",
                                          :dxfId => non_active,
                                          :priority => 1 })
      mailer.add_conditional_formatting("C:#{@letter}", { :type => :cellIs,
                                          :operator => :equal,
                                          :formula => "\"Yes\"",
                                          :dxfId => active,
                                          :priority => 1 })

      
    end
    
    ###PRODUCTS - ALL TAB####
    workbook.add_worksheet(name: "Products - ALL") do |sheet|
      ###Hide gridlines on page
      sheet.sheet_view.show_grid_lines= false
      sheet.sheet_pr.tab_color = @tab1
      
      sheet.add_row []
      sheet.add_row ["","IPA/ISAL and ePacket Product Breakdown"], style: @tab_header
      sheet.add_row ["","Source: PostalOne, " + @title_year + " " + @start_month_name + "-" + @month_name], style: @tab_subheader
      sheet.add_row ["",@pqws].flatten, style: @invisible
      sheet.add_row ["","", "Customer Selector:", "", "All"], :style => @select
      
      ###VOLUME AND REVENUE HEADER PLACEMENT, STYLING, AND MERGING
      sheet.add_row ["", "", @empty_arr, @empty_arr, "Change from Previous Month"].flatten, :style => [nil, nil, @style_nil, @style_nil, @row_header_merge, @row_header_merge, @row_header_merge, @row_header_merge].flatten
      sheet.add_row ["","", @vol_arr, @rev_arr, "Volume Change", "", "Revenue Change" ""].flatten, :style => [nil, nil, @style_merge, @style_merge, @row_header_merge, @row_header_merge, @row_header_merge, @row_header_merge].flatten
      sheet.add_row ["","Product",@date_arr, @date_arr, "Gross", "Percent", "Gross", "Percent"].flatten, :style => [nil, @row_header_right, @date_style, @date_style, @row_header_middle, @row_header_middle, @row_header_middle, @row_header_middle].flatten
      
      @col_arr = ('C'..'Z').to_a
      sheet.merge_cells "C5:D5"
      sheet.merge_cells "#{@col_arr[@rev_arr.length*2]}6:#{@col_arr[@rev_arr.length*2+3]}6"
      sheet.merge_cells "C7:#{@col_arr[@vol_arr.length-1]}7"
      sheet.merge_cells "#{@col_arr[@rev_arr.length]}7:#{@col_arr[(@rev_arr.length*2)-1]}7"
      sheet.merge_cells "#{@col_arr[@rev_arr.length*2]}7:#{@col_arr[@rev_arr.length*2+1]}7"
      sheet.merge_cells "#{@col_arr[@rev_arr.length*2+2]}7:#{@col_arr[@rev_arr.length*2+3]}7"
      
      ###BEGIN ITERATION OF PRODUCT ROWS
      @products.each_with_index do |p, index|
        @temp_arr = [index, p]
        @count = index+8
        
        if p != 'All'
          if ["PMEI", "PMI", "FCPIS"].include?(p)
             @letter_arr = ('F'..'Z').to_a
             
             @vol_last = @letter_arr[@date_arr.length-4] #2
             @vol_recent = @letter_arr[@date_arr.length-5] #1
             @vol_percent = @letter_arr[@date_arr.length*2-3]
             
             @rev_last = @letter_arr[(@date_arr.length*2)-4]
             @rev_recent = @letter_arr[(@date_arr.length*2)-5]
             @rev_percent = @letter_arr[@date_arr.length*2-1]
             
             @num_format = []
            (@date_arr.length*2).times do
              @letter = @letter_arr.shift()
              @temp = %Q|=IF($E$5="All", SUMPRODUCT(('PQW Report Data'!$D$4:$D$3944='Products - ALL'!$B#{index+8})*('PQW Report Data'!#{@letter}$4:#{@letter}$3944)), SUMPRODUCT(('PQW Report Data'!$B$4:$B$3944='Products - ALL'!$E$5)*('PQW Report Data'!$D$4:$D$3944='Products - ALL'!$B#{index+8})*('PQW Report Data'!#{@letter}$4:#{@letter}$3944)))|
              @temp_arr << @temp
            end
             
          else
             @letter_arr = ('E'..'Z').to_a
             
             @vol_last = @letter_arr[@date_arr.length-3] #3 
             @vol_recent = @letter_arr[@date_arr.length-4] #2
             @vol_percent = @letter_arr[@date_arr.length*2-2]
             
             @rev_last = @letter_arr[(@date_arr.length*2)-3] 
             @rev_recent = @letter_arr[(@date_arr.length*2)-4] 
             @rev_percent = @letter_arr[@date_arr.length*2]
             
             (@date_arr.length*2).times do
              @letter = @letter_arr.shift
              @temp = %Q|=IF($E$5="All", SUMPRODUCT(('Postal One PQW Report'!$D$4:$D$9999='Products - ALL'!$B#{index+8})*('Postal One PQW Report'!#{@letter}$4:#{@letter}$9999)), SUMPRODUCT(('Postal One PQW Report'!$D$4:$D$9999='Products - ALL'!$B#{index+8})*('Postal One PQW Report'!$B$4:$B$9999='Products - ALL'!$E$5)*('Postal One PQW Report'!#{@letter}$4:#{@letter}$9999)))|
              @temp_arr << @temp
              end        
             
          end #if pmei, epacket, fcpis end
             if index == (@products.length-1)
              sheet.add_row [@temp_arr,
                             %Q|=IFERROR(#{@vol_last}#{@count}-#{@vol_recent}#{@count}, 0)|,
                             %Q|=IFERROR(#{@vol_percent}#{@count}/#{@vol_last}#{@count}, 0)|,
                             %Q|=IFERROR(#{@rev_last}#{@count}-#{@rev_recent}#{@count}, 0)|,
                             %Q|=IFERROR(#{@rev_percent}#{@count}/#{@rev_last}#{@count}, 0)|].flatten, :style => [@row_num, @last_row_middle, @last_vol, @last_rev, @last_row_vol, @last_row_per, @last_row_vol, @last_row_per].flatten           
             else
              sheet.add_row [@temp_arr,
                             %Q|=IFERROR(#{@vol_last}#{@count}-#{@vol_recent}#{@count}, 0)|,
                             %Q|=IFERROR(#{@vol_percent}#{@count}/#{@vol_last}#{@count}, 0)|,
                             %Q|=IFERROR(#{@rev_last}#{@count}-#{@rev_recent}#{@count}, 0)|,
                             %Q|=IFERROR(#{@rev_percent}#{@count}/#{@rev_last}#{@count}, 0)|].flatten, :style => [@row_num, @row_styling_middle, @vol_style, @rev_style, @row_styling_vol, @row_styling_per, @row_styling_vol, @row_styling_per].flatten
             end
          
        end #if not all end
        
        @final_count = index+1
      end
      
      ###TOTALS ROW###
      @totals = [@final_count+1, "Total"]
      (@date_arr.length*2).times do
        @column = @col_arr.shift
        @temp = "=SUM(#{@column}9:#{@column}22)"
        @totals << @temp
      end
      
      @col_arr = ('C'..'Z').to_a
      @vol_last = @col_arr[@date_arr.length-1]
      @vol_recent = @col_arr[@date_arr.length-2]
      @vol_percent = @col_arr[@date_arr.length*2]
      
      @rev_last = @col_arr[(@date_arr.length*2)-1]
      @rev_recent = @col_arr[(@date_arr.length*2)-2]
      @rev_percent = @col_arr[@date_arr.length*2+2]
      sheet.add_row [@totals,
                     %Q|=IFERROR((#{@vol_last}23-#{@vol_recent}23), 0)|,
                     %Q|=IFERROR(#{@vol_percent}23/#{@vol_last}23, 0)|,
                     %Q|=IFERROR((#{@rev_last}23-#{@rev_recent}23), 0)|,
                     %Q|=IFERROR(#{@rev_percent}23/#{@rev_last}23, 0)|
                    ].flatten, :style => [@row_num, @row_styling_total_text, @vol_style_total, @rev_style_total, @row_styling_vol_total, @row_styling_per_total, @row_styling_vol_total, @row_styling_per_total].flatten
      
      @letter_arr = ('B'..'Z').to_a
      if @pqws.length > @letter_arr.length
        @letter_arr_cont = ('A'..'Z').to_a
        @letter_arr_cont.map! { |word| "A#{word}" }
        @index = @letter_arr_cont[(@pqws.length - @letter_arr.length)-1]
      else
        @index = @letter_arr[@pqws.length]
      end
      
      sheet.add_data_validation("E5", {
      :type => :list,
      :formula1 => "B4:#{@index}4",
      :showDropDown => false,
      :showInputMessage => true,
      :promptTitle => 'Customer',
      :prompt => 'Choose customer to view'
      })
      
      
      
      @format_arr = ('C'..'Z').to_a
      if (@date_arr.length*2) > @format_arr.length
        @format_cont = ('A'..'Z').to_a
        @format_cont.map! { |word| "A#{word}" }
        @format_arr << @format_cont
        @format_arr.flatten!
      end
    
      @last_col = @format_arr[(@date_arr.length*2)+1]
      @second_last_col = @format_arr[(@date_arr.length*2)]
      
      # Apply conditional formatting in the worksheet
      sheet.add_conditional_formatting("#{@second_last_col}9:#{@last_col}23", { :type => :cellIs,
                                          :operator => :lessThan,
                                          :formula => '0',
                                          :dxfId => unprofitable,
                                          :priority => 1 })
      sheet.add_conditional_formatting("#{@second_last_col}9:#{@last_col}23", { :type => :cellIs,
                                          :operator => :greaterThan,
                                          :formula => '0',
                                          :dxfId => profitable,
                                          :priority => 1 })
      
      
      ###FORMAT COLUMNS HARD CODED UNTIL BETTER WAY IS FOUND TO REPLACE IT
      sheet.column_widths(4, 20, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11)
    end
    
    ###CUSTOMERS - ALL TAB####
    workbook.add_worksheet(name: "Customers - ALL") do |sheet|
      ###Hide gridlines on page
      sheet.sheet_view.show_grid_lines= false
      sheet.sheet_pr.tab_color = @tab1
      
      sheet.add_row []
      sheet.add_row ["","Customers, Total Volume Comparison"], style: @tab_header
      sheet.add_row ["","Source: PostalOne & GEPS Database, " + @title_year + " " + @start_month_name + "-" + @month_name], style: @tab_subheader
      sheet.add_row [] 
      
      
      ###VOLUME AND REVENUE HEADER PLACEMENT, STYLING, AND MERGING
      sheet.add_row ["", "", @empty_arr, @empty_arr, "Change from Previous Month"].flatten, :style => [nil, nil, @style_nil, @style_nil, @row_header_merge, @row_header_merge, @row_header_merge, @row_header_merge].flatten
      sheet.add_row ["","", @vol_arr, @rev_arr, "Volume Change", "", "Revenue Change", ""].flatten, :style => [nil, nil, @style_merge, @style_merge, @row_header_merge, @row_header_merge, @row_header_merge, @row_header_merge].flatten
      sheet.add_row ["","Customer",@date_arr, @date_arr, "Gross", "Percent", "Gross", "Percent"].flatten, :style => [nil, @row_header_right, @date_style, @date_style, @row_header_middle, @row_header_middle, @row_header_middle, @row_header_middle].flatten
      
      #MERGING OF COLUMN HEADERS
      @col_arr = ('C'..'Z').to_a
      sheet.merge_cells "#{@col_arr[@rev_arr.length*2]}5:#{@col_arr[@rev_arr.length*2+3]}5"
      sheet.merge_cells "C6:#{@col_arr[@vol_arr.length-1]}6"
      sheet.merge_cells "#{@col_arr[@rev_arr.length]}6:#{@col_arr[@rev_arr.length*2-1]}6"
      sheet.merge_cells "#{@col_arr[@rev_arr.length*2]}6:#{@col_arr[@rev_arr.length*2+1]}6"
      sheet.merge_cells "#{@col_arr[@rev_arr.length*2+2]}6:#{@col_arr[@rev_arr.length*2+3]}6"
      
      ###BEGIN ITERATION OF CUSTOMER ROWS
      @pqws.each_with_index do |p, index|
        if p != "All"
          
          @temp_arr = [index, p]
          @count = index+7
              
             #LETTER ARRAY REFERENCING COLUMNS
             @letter_arr = ('C'..'Z').to_a
             @p1_letter_arr = ('E'..'Z').to_a
             @pqw_letter_arr = ('F'..'Z').to_a
             
             #SETTING VOLUME COMPARISON COLUMN POSITIONS
             @vol_last = @letter_arr[@date_arr.length-1]
             @vol_recent = @letter_arr[@date_arr.length-2]
             @vol_percent = @letter_arr[@date_arr.length*2]
             
             #SETTING REVENUE COMPARISON COLUMN POSITIONS
             @rev_last = @letter_arr[(@date_arr.length*2)-1]
             @rev_recent = @letter_arr[(@date_arr.length*2)-2]
             @rev_percent = @letter_arr[@date_arr.length*2+2]
             
             #RUN THROUGH PQWs, CREATING ROWS WITH FORMULAS
            (@date_arr.length*2).times do
              @p1_letter = @p1_letter_arr.shift()
              @pqw_letter = @pqw_letter_arr.shift()
              @temp = %Q|=SUMPRODUCT(('Postal One PQW Report'!$B$4:$B$9999='Customers - ALL'!$B#{index+7})*('Postal One PQW Report'!#{@p1_letter}$4:#{@p1_letter}$9999))+SUMPRODUCT(('PQW Report Data'!$B$4:$B$3944='Customers - ALL'!$B#{index+7})*('PQW Report Data'!#{@pqw_letter}$4:#{@pqw_letter}$3944))|
                      
              @temp_arr << @temp
             end
             
             #ADD ROW to SHEET
             if index == (@pqws.length-1)
              sheet.add_row [@temp_arr,
                             %Q|=IFERROR(#{@vol_last}#{@count}-#{@vol_recent}#{@count}, 0)|,
                             %Q|=IFERROR(#{@vol_percent}#{@count}/#{@vol_last}#{@count}, 0)|,
                             %Q|=IFERROR(#{@rev_last}#{@count}-#{@rev_recent}#{@count}, 0)|,
                             %Q|=IFERROR(#{@rev_percent}#{@count}/#{@rev_last}#{@count}, 0)|].flatten, :style => [@row_num, @last_row_middle, @last_vol, @last_rev, @last_row_vol, @last_row_per, @last_row_vol, @last_row_per].flatten           
             else
              sheet.add_row [@temp_arr,
                             %Q|=IFERROR(#{@vol_last}#{@count}-#{@vol_recent}#{@count}, 0)|,
                             %Q|=IFERROR(#{@vol_percent}#{@count}/#{@vol_last}#{@count}, 0)|,
                             %Q|=IFERROR(#{@rev_last}#{@count}-#{@rev_recent}#{@count}, 0)|,
                             %Q|=IFERROR(#{@rev_percent}#{@count}/#{@rev_last}#{@count}, 0)|].flatten, :style => [@row_num, @row_styling_middle, @vol_style, @rev_style, @row_styling_vol, @row_styling_per, @row_styling_vol, @row_styling_per].flatten
             end
          @final_count = index
        end
      end
      
      ###TOTALS ROW###
      @totals = [@final_count+1, "Total"]
      (@date_arr.length*2).times do
        @column = @col_arr.shift
        @temp = "=SUM(#{@column}9:#{@column}#{(@pqws.length-1)+7})"
        @totals << @temp
      end
      
      #ADD TOTALS ROW TO SHEET
      sheet.add_row [@totals,
                     %Q|=IFERROR(#{@vol_last}#{@final_count+8}-#{@vol_recent}#{@final_count+8}, 0)|,
                     %Q|=IFERROR(#{@vol_percent}#{@final_count+8}/#{@vol_last}#{@final_count+8}, 0)|,
                     %Q|=IFERROR(#{@rev_last}#{@final_count+8}-#{@rev_recent}#{@final_count+8}, 0)|,
                     %Q|=IFERROR(#{@rev_percent}#{@final_count+8}/#{@rev_last}#{@final_count+8}, 0)|
                    ].flatten, :style => [@row_num, @row_styling_total_text, @vol_style_total, @rev_style_total, @row_styling_vol_total, @row_styling_per_total, @row_styling_vol_total, @row_styling_per_total].flatten
      
      @conditional_last_col = @format_arr[(@date_arr.length*2)+3]
      @last_col = @pqws.length+9
      # Apply conditional formatting in the worksheet
      sheet.add_conditional_formatting("#{@second_last_col}8:#{@conditional_last_col}#{@last_col}", { :type => :cellIs,
                                          :operator => :lessThan,
                                          :formula => '0',
                                          :dxfId => unprofitable,
                                          :priority => 1 })
      sheet.add_conditional_formatting("#{@second_last_col}8:#{@conditional_last_col}#{@last_col}", { :type => :cellIs,
                                          :operator => :greaterThan,
                                          :formula => '0',
                                          :dxfId => profitable,
                                          :priority => 1 })
      
      
      ###FORMAT COLUMNS HARD CODED UNTIL BETTER WAY IS FOUND TO REPLACE IT
      sheet.column_widths(4, 25, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12)
    end
    
    ###POSTAL ONE RATEGROUP BREAKDOWN####
    workbook.add_worksheet(name: "Postal One Rate Group Breakdown") do |sheet|
      ###Hide gridlines on page
      sheet.sheet_view.show_grid_lines= false
      sheet.sheet_pr.tab_color = @tab2
  
      sheet.add_row []
      sheet.add_row ["","IPA/ISAL and ePacket Rategroup Breakdown"], style: @tab_header
      sheet.add_row ["","Source: PostalOne, " + @title_year + " " + @start_month_name + "-" + @month_name], style: @tab_subheader
      sheet.add_row ["",@pqws].flatten, style: @invisible
      sheet.add_row ["","", "Customer Selector:", "", "All"], :style => @select
      
      ###VOLUME AND REVENUE HEADER PLACEMENT, STYLING, AND MERGING
      sheet.add_row ["","", @empty_arr, @empty_arr, "Change from Previous Month"].flatten, :style => [nil, nil, @style_nil, @style_nil, @row_header_merge, @row_header_merge, @row_header_merge, @row_header_merge].flatten
      sheet.add_row ["","", @vol_arr, @rev_arr, "Volume", "", "Revenue", ""].flatten, :style => [nil, nil, @style_merge, @style_merge, @row_header_merge, @row_header_merge, @row_header_merge, @row_header_merge].flatten
      sheet.add_row ["","Product",@date_arr, @date_arr, "Gross", "Percent", "Gross", "Percent"].flatten, :style => [nil, @row_header_right, @date_style, @date_style, @row_header_middle, @row_header_middle, @row_header_middle, @row_header_middle].flatten
      
      @col_arr = ('C'..'Z').to_a
      sheet.merge_cells "C5:D5"
      sheet.merge_cells "#{@col_arr[@rev_arr.length*2]}6:#{@col_arr[@rev_arr.length*2+3]}6"
      sheet.merge_cells "C7:#{@col_arr[@vol_arr.length-1]}7"
      sheet.merge_cells "#{@col_arr[@rev_arr.length]}7:#{@col_arr[(@rev_arr.length*2)-1]}7"
      sheet.merge_cells "#{@col_arr[@rev_arr.length*2]}7:#{@col_arr[@rev_arr.length*2+1]}7"
      sheet.merge_cells "#{@col_arr[@rev_arr.length*2+2]}7:#{@col_arr[@rev_arr.length*2+3]}7"
      
      ###BEGIN ITERATION OF RATEGROUP ROWS
      @postal_one_rategroups.each_with_index do |r, index|
        @temp_arr = [index, r]
        @count = index+8
         if r != 'All'
           @letter_arr = ('E'..'Z').to_a
           
           @vol_last = @letter_arr[@date_arr.length-3] #3
           @vol_recent = @letter_arr[@date_arr.length-4] #2
           @vol_percent = @letter_arr[@date_arr.length*2-2]
             
           @rev_last = @letter_arr[(@date_arr.length*2)-3]
           @rev_recent = @letter_arr[(@date_arr.length*2)-4] 
           @rev_percent = @letter_arr[@date_arr.length*2]
           
           
           (@date_arr.length*2).times do
            @letter = @letter_arr.shift
            @temp = %Q|=IF($E$5="All",SUMPRODUCT( -- ('Postal One PQW Report'!$C$3:$C$2386=$B#{index+8}),'Postal One PQW Report'!#{@letter}$3:#{@letter}$2386),  SUMPRODUCT(('Postal One PQW Report'!$C$4:$C$2126=$B#{index+8})*('Postal One PQW Report'!$B$4:$B$2126=UPPER($E$5))*('Postal One PQW Report'!#{@letter}$4:#{@letter}$2126)))|
            @temp_arr << @temp
           end
           
           if index == (@postal_one_rategroups.length-1)
            sheet.add_row [@temp_arr,
                           %Q|=IFERROR(#{@vol_last}#{@count}-#{@vol_recent}#{@count}, 0)|,
                           %Q|=IFERROR(#{@vol_percent}#{@count}/#{@vol_last}#{@count}, 0)|,
                           %Q|=IFERROR(#{@rev_last}#{@count}-#{@rev_recent}#{@count}, 0)|,
                           %Q|=IFERROR(#{@rev_percent}#{@count}/#{@rev_last}#{@count}, 0)|].flatten, :style => [@row_num, @last_row_middle, @last_vol, @last_rev, @last_row_vol, @last_row_per, @last_row_vol, @last_row_per].flatten           
           else
            sheet.add_row [@temp_arr,
                           %Q|=IFERROR(#{@vol_last}#{@count}-#{@vol_recent}#{@count}, 0)|,
                           %Q|=IFERROR(#{@vol_percent}#{@count}/#{@vol_last}#{@count}, 0)|,
                           %Q|=IFERROR(#{@rev_last}#{@count}-#{@rev_recent}#{@count}, 0)|,
                           %Q|=IFERROR(#{@rev_percent}#{@count}/#{@rev_last}#{@count}, 0)|].flatten, :style => [@row_num, @row_styling_middle, @vol_style, @rev_style, @row_styling_vol, @row_styling_per, @row_styling_vol, @row_styling_per].flatten
           end
         end
        @final_count = index+1
      end
      
      ###TOTALS ROW###
      @totals = [@final_count+1, "Total"]
      (@date_arr.length*2).times do
        @column = @col_arr.shift
        @temp = "=SUM(#{@column}9:#{@column}#{(@postal_one_rategroups.length-1)+8})"
        @totals << @temp
      end

      @col_arr = ('C'..'Z').to_a
      @vol_last = @col_arr[@date_arr.length-1]
      @vol_recent = @col_arr[@date_arr.length-2]
      @vol_percent = @col_arr[@date_arr.length*2]
      
      @rev_last = @col_arr[(@date_arr.length*2)-1]
      @rev_recent = @col_arr[(@date_arr.length*2)-2]
      @rev_percent = @col_arr[@date_arr.length*2+2]
      
      @row_index = @postal_one_rategroups.length+9
      sheet.add_row [@totals,
                     %Q|=IFERROR(#{@vol_last}#{@row_index}-#{@vol_recent}#{@row_index}, 0)|,
                     %Q|=IFERROR(#{@vol_percent}#{@row_index}/#{@vol_last}#{@row_index}, 0)|,
                     %Q|=IFERROR(#{@rev_last}#{@row_index}-#{@rev_recent}#{@row_index}, 0)|,
                     %Q|=IFERROR(#{@rev_percent}#{@row_index}/#{@rev_last}#{@row_index}, 0)|
                    ].flatten, :style => [@row_num, @row_styling_total_text, @vol_style_total, @rev_style_total, @row_styling_vol_total, @row_styling_per_total, @row_styling_vol_total, @row_styling_per_total].flatten
          
      @letter_arr = ('B'..'Z').to_a
      if @pqws.length > @letter_arr.length
        @letter_arr_cont = ('A'..'Z').to_a
        @letter_arr_cont.map! { |word| "A#{word}" }
        @index = @letter_arr_cont[(@pqws.length - @letter_arr.length)-1]
        
      else
        @index = @letter_arr[@pqws.length]
          
      end
      
      sheet.add_data_validation("E5", {
      :type => :list,
      :formula1 => "B4:#{@index}4",
      :showDropDown => false,
      :showInputMessage => true,
      :promptTitle => 'Customer',
      :prompt => 'Choose customer to view (Default: )'
      })
      
      #Conditoinal format arr
     @format_arr = ('C'..'Z').to_a
      if (@date_arr.length*2) > @format_arr.length
        @format_cont = ('A'..'Z').to_a
        @format_cont.map! { |word| "A#{word}" }
        @format_arr << @format_cont
        @format_arr.flatten!
      end
    
      @last_col = @format_arr[(@date_arr.length*2)+1]
      @second_last_col = @format_arr[(@date_arr.length*2)]
      
      # Apply conditional formatting in the worksheet
      sheet.add_conditional_formatting("#{@second_last_col}9:#{@last_col}29", { :type => :cellIs,
                                          :operator => :lessThan,
                                          :formula => '0',
                                          :dxfId => unprofitable,
                                          :priority => 1 })
      sheet.add_conditional_formatting("#{@second_last_col}9:#{@last_col}29", { :type => :cellIs,
                                          :operator => :greaterThan,
                                          :formula => '0',
                                          :dxfId => profitable,
                                          :priority => 1 })
      
      ###FORMAT COLUMNS HARD CODED UNTIL BETTER WAY IS FOUND TO REPLACE IT
      sheet.column_widths(4, 20, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10)
    end    
    
    
    ###GEPS VOLUME & DECLINES###
    workbook.add_worksheet(name: "GEPS Volume & Declines") do |sheet|
      
      ###Hide gridlines on page
      sheet.sheet_view.show_grid_lines= false
      sheet.sheet_pr.tab_color = @tab3
      
      sheet.add_row []
      sheet.add_row ["", "GEPS Volume Dashboard", "", "", "", "", "", "", "", "", "", "Top GEPS Declines "], :style => [nil, @tab_header, nil, nil, nil, nil, nil, nil, nil, nil, nil, @tab_header]
      sheet.add_row ["", "Source: GEPS/NMATS, " + @title_year + " " + @start_month_name + "-" + @month_name, "", "", "", "", "", "", "", "", "", "For Month of " + @month_name], :style => [nil, @tab_subheader, nil, nil, nil, nil, nil, nil, nil, nil, nil, @tab_subheader]
      sheet.add_row ["" ,"Customer" ,"All", "Product:", "All", "Rategroup:", "All", "Weightstep:", "All"], :style => @select
      sheet.add_row []
      sheet.add_row ["", "Calendar Month", "Total Volume", "", "", "", "", "", "", "", "", "", "Mailer", "Rategroup", "Product", "Weightstep", "Volume Change"], :style => [nil, @row_header, @row_header, nil, nil, nil, nil, nil, nil, nil, nil, nil, @row_header, @row_header, @row_header, @row_header, @row_header]
      
      
      @letter_arr = ('F'..'Z').to_a
      
      12.times do |count|
        @display_arr = []
        if count <= (@date_arr.length-1)
          @letter = @letter_arr.shift()
          @date = Date::MONTHNAMES[@date_arr[count].to_i]
          @temp = %Q|=IF(AND($C$4="All",$E$4="All",$G$4="All",$I$4="All"),SUMPRODUCT('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340),
                      IF(AND($C$4="All",$E$4="All",$G$4="All"),SUMPRODUCT(--('PQW Report Data'!$E$4:$E$89340='GEPS Volume & Declines'!$I$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340)),
                      IF(AND($C$4="All",$G$4="All",$I$4="All"),SUMPRODUCT(--('PQW Report Data'!$D$4:$D$89340='GEPS Volume & Declines'!$E$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340)),
                      IF(AND($C$4="All",$E$4="All",$I$4="All"),SUMPRODUCT(--('PQW Report Data'!$C$4:$C$89340='GEPS Volume & Declines'!GI$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340)),
                      IF(AND($E$4="All",$G$4="All",$I$4="All"),SUMPRODUCT(--('PQW Report Data'!$B$4:$B$89340='GEPS Volume & Declines'!$C$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340)),
                      IF(AND($C$4="All",$E$4="All"),SUMPRODUCT(--('PQW Report Data'!$C$4:$C$89340='GEPS Volume & Declines'!$G$4),--('PQW Report Data'!$E$4:$E$89340='GEPS Volume & Declines'!$I$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340)),
                      IF(AND($C$4="All",$G$4="All"),SUMPRODUCT(--('PQW Report Data'!$D$4:$D$89340='GEPS Volume & Declines'!$E$4),--('PQW Report Data'!$E$4:$E$89340='GEPS Volume & Declines'!$I$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340)),
                      IF(AND($C$4="All",$I$4="All"),SUMPRODUCT(--('PQW Report Data'!$D$4:$D$89340='GEPS Volume & Declines'!$E$4),--('PQW Report Data'!$C$4:$C$89340='GEPS Volume & Declines'!$G$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340)),
                      IF(AND($E$4="All",$G$4="All"),SUMPRODUCT(--('PQW Report Data'!$B$4:$B$89340='GEPS Volume & Declines'!$C$4),--('PQW Report Data'!$E$4:$E$89340='GEPS Volume & Declines'!$I$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340)),
                      IF(AND($E$4="All",$I$4="All"),SUMPRODUCT(--('PQW Report Data'!$B$4:$B$89340='GEPS Volume & Declines'!$C$4),--('PQW Report Data'!$C$4:$C$89340='GEPS Volume & Declines'!$G$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340)),
                      IF(AND($G$4="All",$I$4="All"),SUMPRODUCT(--('PQW Report Data'!$B$4:$B$89340='GEPS Volume & Declines'!$C$4),--('PQW Report Data'!$D$4:$D$89340='GEPS Volume & Declines'!$E$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340)),
                      IF($C$4="All",SUMPRODUCT(--('PQW Report Data'!$D$4:$D$89340='GEPS Volume & Declines'!$E$4),--('PQW Report Data'!$C$4:$C$89340='GEPS Volume & Declines'!$G$4),--('PQW Report Data'!$E$4:$E$89340='GEPS Volume & Declines'!$I$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340)),
                      IF($E$4="All",SUMPRODUCT(--('PQW Report Data'!$B$4:$B$89340='GEPS Volume & Declines'!$C$4),--('PQW Report Data'!$C$4:$C$89340='GEPS Volume & Declines'!$G$4),--('PQW Report Data'!$E$4:$E$89340='GEPS Volume & Declines'!$I$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340)),
                      IF($G$4="All",SUMPRODUCT(--('PQW Report Data'!$B$4:$B$89340='GEPS Volume & Declines'!$C$4),--('PQW Report Data'!$D$4:$D$89340='GEPS Volume & Declines'!$E$4),--('PQW Report Data'!$E$4:$E$89340='GEPS Volume & Declines'!$I$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340)),
                      IF($I$4="All",SUMPRODUCT(--('PQW Report Data'!$B$4:$B$89340='GEPS Volume & Declines'!$C$4),--('PQW Report Data'!$D$4:$D$89340='GEPS Volume & Declines'!$E$4),--('PQW Report Data'!$C$4:$C$89340='GEPS Volume & Declines'!$G$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340)),
                      SUMPRODUCT(--('PQW Report Data'!$B$4:$B$89340='GEPS Volume & Declines'!$C$4),--('PQW Report Data'!$D$4:$D$89340='GEPS Volume & Declines'!$E$4),--('PQW Report Data'!$C$4:$C$89340='GEPS Volume & Declines'!$G$4),--('PQW Report Data'!$E$4:$E$89340='GEPS Volume & Declines'!$I$4),('PQW Report Data'!$#{@letter}$4:$#{@letter}$89340))))
                      )))))))))))))|
          @display_arr << [count+1, @date, @temp]
        else
          @display_arr << ["", "", ""]
        end
        
        @space_diff = ["", "", "", "", "", "", "", ""]
        @declines = [count+1, %Q|=UPPER('Top Granular Level Declines'!B#{count+5})|,
                     %Q|='Top Granular Level Declines'!C#{count+5}|,
                     %Q|='Top Granular Level Declines'!D#{count+5}|,
                     %Q|='Top Granular Level Declines'!E#{count+5}|,
                     %Q|='Top Granular Level Declines'!F#{count+5}|
                    ]
        @display_arr << [@space_diff, @declines]
        @style_test = @display_arr.flatten
        if @style_test[0] == ""
          @beginning = [nil, nil, nil]
        else
          @beginning = [@row_num, @row_styling_middle, @row_styling_vol]
        end
        
        @space_arr = []
        @space_diff.length.times do
          @temp = nil
          @space_arr << @temp
        end
        
          sheet.add_row [@display_arr].flatten, :style => [@beginning, @space_arr, @row_num, @row_styling_middle, @row_styling_middle, @row_styling_middle, @row_styling_middle, @row_styling_middle].flatten
        
      end
      
      #style specific columns
      sheet.col_style 2, @row_styling_vol, row_offset: 7
      sheet.col_style 16, @row_styling_vol, row_offset: 7
      
      #give more space to headers
      sheet.merge_cells "B2:C2"
      sheet.merge_cells "B3:E3"
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
      sheet.add_row [@geps_pqws].flatten, :style => @invisible
      sheet.add_row [@products].flatten, :style => @invisible
      sheet.add_row [@geps_rategroups].flatten, :style => @invisible
      sheet.add_row [@weightsteps].flatten, :style => @invisible
      
      @col_arr = ('A'..'Z').to_a
      
      sheet.add_data_validation("C4", {
      :type => :list,
      :formula1 => "A19:#{@col_arr[(@geps_pqws.length-1)]}19",
      :showDropDown => false,
      :showInputMessage => true,
      :promptTitle => 'Customer',
      :prompt => 'Choose customer to view'
      })
      
      sheet.add_data_validation("E4", {
      :type => :list,
      :formula1 => "A20:O20",
      :showDropDown => false,
      :showInputMessage => true,
      :promptTitle => 'Product',
      :prompt => 'Choose Product to view'
      })
      
      sheet.add_data_validation("G4", {
      :type => :list,
      :formula1 => "A21:Z21",
      :showDropDown => false,
      :showInputMessage => true,
      :promptTitle => 'Rategroup',
      :prompt => 'Choose rate group to view'
      })
      
      sheet.add_data_validation("I4", {
      :type => :list,
      :formula1 => "A22:BT22",
      :showDropDown => false,
      :showInputMessage => true,
      :promptTitle => 'Weightstep',
      :prompt => 'Choose weightstep to view'
      })
      
      #LINE CHART
      chart = sheet.add_chart(Axlsx::LineChart, :start_at=> "B19", :end_at=> "K30", :show_legend => false, :title=>"Monthly Volume")
            chart.add_series :data => sheet["C7:C#{@date_arr.length+6}"], :labels => sheet["B7:B#{@date_arr.length+6}"], :show_marker => true
            chart.valAxis.gridlines = false
            chart.catAxis.gridlines = false
      
      sheet.column_widths 5, 15, 18, 15, 18, 15, 18, 15, 18, 2, 2, 5, 15, 15, 15, 15, 15
    end
    
    
    
    ###GEPS CHARTS###
    workbook.add_worksheet(name: "GEPS Charts") do |sheet|
      ###Hide gridlines on page
      sheet.sheet_view.show_grid_lines= false
      sheet.sheet_pr.tab_color = @tab3
      
      sheet.add_row []
      sheet.add_row ["", "GEPS Charts"], :style => @tab_header
      sheet.add_row ["", "Source: GEPS/NMATS, " + @title_year + " " + @start_month_name + "-" + @month_name], :style => @tab_subheader
      sheet.add_row ["", "Selected Customer*", "='GEPS Volume & Declines'!C4", "Selected Product*", "='GEPS Volume & Declines'!E4", "", "*Alter these choices by using the dropdowns on the 'GEPS Volume & Declines' tab"], :style => @wrap_text
      sheet.add_row []
      sheet.add_row []
      sheet.add_row []
      
      sheet.merge_cells "G4:O4"
      
      
            @letter_arr = ('B'..'Z').to_a
            @three_back = @letter_arr[@date_arr.length-2]
            @two_back = @letter_arr[@date_arr.length-1]
            @one = @letter_arr[@date_arr.length]
            
            
            @three_mon = @end_date << 2
            @two_mon = @end_date << 1
            
            @three_month = Date::MONTHNAMES[@three_mon.month]
            @two_month = Date::MONTHNAMES[@two_mon.month]
            @one_month = Date::MONTHNAMES[@end_month]
            
            sheet.add_row ["","#{@three_month}", "#{@two_month}", "#{@one_month}"]
            
            16.times do |count|
              sheet.add_row ["='Monthly Volume by Weight Step'!B#{count+10}",
                             "='Monthly Volume by Weight Step'!#{@three_back}#{count+10}",
                             "='Monthly Volume by Weight Step'!#{@two_back}#{count+10}",
                             "='Monthly Volume by Weight Step'!#{@one}#{count+10}",
                             "='Monthly Volume by WS & RG'!AB#{count+10}",
                             "='Monthly Volume Change by WS & RG'!AB#{count+10}"], :style => @invisible
            end
            
            @geps_rategroups.each_with_index do |g, index|
              sheet.add_row [g, "='Monthly Volume by Rate Group'!#{@three_back}#{index+9}",
                                "='Monthly Volume by Rate Group'!#{@two_back}#{index+9}",
                                "='Monthly Volume by Rate Group'!#{@one}#{index+9}"], :style => @invisible
            end
            
            
            sheet.add_chart(Axlsx::Bar3DChart, :start_at => "B6", :end_at => "J26", :title=> "Top Weight Step (0.5-15) Volume Month over Month", :barDir => :col, :legend_position => :b) do |bars|
              bars.add_series :data => sheet["B9:B24"], :labels => sheet["A9:A24"], :title => sheet["B8"]
              bars.add_series :data => sheet["C9:C24"], :labels => sheet["A9:A24"], :title => sheet["C8"]
              bars.add_series :data => sheet["D9:D24"], :labels => sheet["A9:A24"], :title => sheet["D8"]
              bars.valAxis.gridlines = false
              bars.catAxis.gridlines = false
            end
            
            sheet.add_chart(Axlsx::Bar3DChart, :start_at => "K6", :end_at => "T26", :title=> "Total Volume By Rate Group Month over Month", :barDir => :col, :legend_position => :b) do |rategroup|
              rategroup.add_series :data => sheet["B25:B49"], :labels => sheet["A25:A49"], :title => sheet["B8"]
              rategroup.add_series :data => sheet["C25:C49"], :labels => sheet["A25:A49"], :title => sheet["C8"]
              rategroup.add_series :data => sheet["D25:D49"], :labels => sheet["A25:A49"], :title => sheet["D8"]
              rategroup.valAxis.gridlines = false
              rategroup.catAxis.gridlines = false
            end
            
            sheet.add_chart(Axlsx::Bar3DChart, :start_at => "B27", :end_at => "J47", :title=> "Total PQW Volume by Weightsteps 0.5-15", :barDir => :col, :show_legend => false) do |pqws|
              pqws.add_series :data => sheet["E9:E24"], :labels => sheet["A9:A24"]
              pqws.valAxis.gridlines = false
              pqws.catAxis.gridlines = false
            end
            
            sheet.add_chart(Axlsx::Bar3DChart, :start_at => "K27", :end_at => "T47", :title=> "Month over Month Volume Change by Weightsteps 0.5-15", :barDir => :col, :show_legend => false) do |change|
              change.add_series :data => sheet["F9:F24"], :labels => sheet["A9:A24"]
              change.valAxis.gridlines = false
              change.catAxis.gridlines = false
            end
            
        sheet.column_widths 5,15,18,15,18,5,5,5,5
    end
    
    
    ###WEIGHTSTEPS MONTH OVER MONTH###
    workbook.add_worksheet(name: "Monthly Volume by Weight Step") do |sheet|
      
      ###Hide gridlines on page
      sheet.sheet_view.show_grid_lines= false
      sheet.sheet_pr.tab_color = @tab3
      
      sheet.add_row []
      sheet.add_row ["", "Monthly Volume by Weight Step"], :style => @tab_header
      sheet.add_row ["", "Includes Volume from all Rategroups"], :style => @tab_subheader
      sheet.add_row ["", "Source: GEPS/NMATS, "+ @title_year + " " + @start_month_name + "-" + @month_name], :style => @tab_subheader
      sheet.add_row []
      sheet.add_row ["","","Selected Customer*","='GEPS Volume & Declines'!$C$4","Selected Product*","='GEPS Volume & Declines'!$E$4", "", "*Alter these choices by using the dropdowns on the 'GEPS Volume & Declines' tab"], :style => @wrap_text
      sheet.add_row []
      
      sheet.merge_cells "H6:L6"
      @header_arr = ["Calendar Month"]
      (@date_arr.length-1).times do
        @temp = ""
        @header_arr << @temp
      end
      
      sheet.add_row ["", "", @header_arr, "Prior Month Change", ""].flatten, :style => [nil, @row_header, @style_merge, @row_header_merge, @row_header_merge].flatten
      sheet.add_row ["", "Weight Step", @date_arr, "Gross Change", "Percent Change"].flatten, :style => [nil, @row_header, @date_style, @row_header_middle, @row_header_middle].flatten
      
      
      @weightsteps.each_with_index do |w, index|
        
        if w != 'All'
          @row_arr = ["", w] 
          @letter_arr = ('F'..'Z').to_a
          (@date_arr.length).times do 
            @letter = @letter_arr.shift()
            @temp = %Q|=IF(AND($D$6="All",$F$6="All"),SUMPRODUCT(('PQW Report Data'!$E$4:$E$11233=$B#{index+9})*('PQW Report Data'!#{@letter}$4:#{@letter}$11233)),
                     IF($D$6="All", SUMPRODUCT(('PQW Report Data'!$E$4:$E$11233=$B#{index+9})*('PQW Report Data'!$D$4:$D$11233='GEPS Report Dashboard & Charts'!$E$4)*('PQW Report Data'!#{@letter}$4:#{@letter}$11233)),
                     IF($F$6="All",  SUMPRODUCT(('PQW Report Data'!$E$4:$E$11233=$B#{index+9})*('PQW Report Data'!$B$4:$B$11233='GEPS Report Dashboard & Charts'!$C$4)*('PQW Report Data'!#{@letter}$4:#{@letter}$11233)),
                     SUMPRODUCT(('PQW Report Data'!$E$4:$E$11233=$B#{index+9})*('PQW Report Data'!$B$4:$B$11233='GEPS VOlume & Declines'!$C$4)*('PQW Report Data'!$D$4:$D$11233='GEPS VOlume & Declines'!$E$4)*('PQW Report Data'!#{@letter}$4:#{@letter}$11233)))))|
            @row_arr << @temp
          end
          
          @col_arr = ('B'..'Z').to_a
          
          @row_arr << ["=#{@col_arr[@date_arr.length]}#{index+9}-#{@col_arr[@date_arr.length-1]}#{index+9}", "=IFERROR(#{@col_arr[@date_arr.length+1]}#{index+9}/#{@col_arr[@date_arr.length-1]}#{index+9}, 0)"]
          
          sheet.add_row [@row_arr].flatten, :style =>[nil, @row_styling_middle, @vol_style, @row_styling_vol, @row_styling_per].flatten
        end
      end    
      
      @col_arr = ('C'..'Z').to_a
      sheet.merge_cells "C8:#{@col_arr[@date_arr.length-1]}8"
      sheet.merge_cells "#{@col_arr[@date_arr.length]}8:#{@col_arr[@date_arr.length+1]}8"
      
      # Apply conditional formatting in the worksheet
      sheet.add_conditional_formatting("#{@col_arr[@date_arr.length]}10:#{@col_arr[@date_arr.length+1]}80", { :type => :cellIs,
                                          :operator => :lessThan,
                                          :formula => '0',
                                          :dxfId => unprofitable,
                                          :priority => 1 })
      sheet.add_conditional_formatting("#{@col_arr[@date_arr.length]}10:#{@col_arr[@date_arr.length+1]}80", { :type => :cellIs,
                                          :operator => :greaterThan,
                                          :formula => '0',
                                          :dxfId => profitable,
                                          :priority => 1 })
      
      sheet.column_widths 5, 11, 11, 16, 11, 16, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11
    end
    
    ###RATEGROUPS MONTH OVER MONTH###
    workbook.add_worksheet(name: "Monthly Volume by Rate Group") do |sheet|
      
      ###Hide gridlines on page
      sheet.sheet_view.show_grid_lines= false
      sheet.sheet_pr.tab_color = @tab3
      
      sheet.add_row []
      sheet.add_row ["","Monthly Volume by Rate Group"], :style => @tab_header
      sheet.add_row ["", "Includes Volume from All Weightsteps"], :style => @tab_subheader
      sheet.add_row ["", "Source :GEPS/NMATS, " + @title_year + " " + @start_month_name + "-" + @month_name], :style => @tab_subheader
      sheet.add_row ["","","Selected Customer*","='GEPS Volume & Declines'!$C$4","Selected Product*","='GEPS Volume & Declines'!$E$4", "", "*Alter these choices by using the dropdowns on the 'GEPS Volume & Declines' tab"], :style => @wrap_text
      sheet.add_row []
      sheet.merge_cells "H5:L5"
      
      @header = ["", "", "Calendar Month"]
      (@date_arr.length-1).times do
        @temp = ""
        @header << @temp
      end
      
      sheet.add_row [@header, "Prior Month Change", ""].flatten, :style => [nil, @row_header, @style_merge, @row_header_merge, @row_header_merge].flatten
      sheet.add_row ["", "Rategroup", @date_arr, "Gross Change", "Percent Change"].flatten, :style => [nil, @row_header, @date_style, @row_header_middle, @row_header_middle].flatten
      
      @geps_rategroups.each_with_index do |r, index|
        if r != 'All'
          @temp_arr = ["", r]
          @letter_arr = ('F'..'Z').to_a
          
           (@date_arr.length).times do
              @letter = @letter_arr.shift()
              @temp = %Q|=IF(AND($D$5="All",$F$5="All"),SUMPRODUCT(('PQW Report Data'!$C$4:$C$11233=$B#{index+8})*('PQW Report Data'!#{@letter}$4:#{@letter}$11233)),
                      IF($D$5="All", SUMPRODUCT(('PQW Report Data'!$C$4:$C$11233=$B#{index+8})*('PQW Report Data'!$D$4:$D$11233='GEPS Volume & Declines'!$C$4)*('PQW Report Data'!#{@letter}$4:#{@letter}$11233)),
                      IF($F$5="All",  SUMPRODUCT(('PQW Report Data'!$C$4:$C$11233=$B#{index+8})*('PQW Report Data'!$B$4:$B$11233='GEPS Volume & Declines'!$C$4)*('PQW Report Data'!#{@letter}$4:#{@letter}$11233)),SUMPRODUCT(('PQW Report Data'!$C$4:$C$11233=$B#{index+8})*('PQW Report Data'!$B$4:$B$11233='GEPS Volume & Declines'!$C$4)*('PQW Report Data'!$D$4:$D$11233='GEPS Report Dashboard & Charts'!$C$4)*('PQW Report Data'!#{@letter}$4:#{@letter}$11233)))))|
              @temp_arr << @temp
            end
            
            @col_arr = ('B'..'Z').to_a
            
            @temp_arr << ["=#{@col_arr[@date_arr.length]}#{index+8}-#{@col_arr[@date_arr.length-1]}#{index+8}", "=IFERROR(#{@col_arr[@date_arr.length+1]}#{index+8}/#{@col_arr[@date_arr.length-1]}#{index+8}, 0)"]
            
            sheet.add_row [@temp_arr].flatten, :style => [nil, @row_styling_middle, @vol_style, @row_styling_vol, @row_styling_per].flatten
        end
      end
      
      
      @form_arr = ('C'..'Z').to_a
      # Apply conditional formatting in the worksheet
      sheet.add_conditional_formatting("#{@form_arr[@date_arr.length]}9:#{@form_arr[@date_arr.length+1]}33", { :type => :cellIs,
                                          :operator => :lessThan,
                                          :formula => '0',
                                          :dxfId => unprofitable,
                                          :priority => 1 })
      sheet.add_conditional_formatting("#{@form_arr[@date_arr.length]}9:#{@form_arr[@date_arr.length+1]}33", { :type => :cellIs,
                                          :operator => :greaterThan,
                                          :formula => '0',
                                          :dxfId => profitable,
                                          :priority => 1 })
      
      @col_arr = ('C'..'Z').to_a
      sheet.merge_cells "C7:#{@col_arr[@date_arr.length-1]}7"
      sheet.merge_cells "#{@col_arr[@date_arr.length]}7:#{@col_arr[@date_arr.length+1]}7"
      
      sheet.column_widths 5, 8, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11
      
    end
    
    ###TOTAL VOLUME BY WS&RG###
    workbook.add_worksheet(name: "Monthly Volume by WS & RG") do |sheet|
      
      ###Hide gridlines on page
      sheet.sheet_view.show_grid_lines= false
      sheet.sheet_pr.tab_color = @tab3
      
      sheet.add_row []
      sheet.add_row ["", "Monthly Volume by WS & RG"], :style => @tab_header
      sheet.add_row ["", "Source: GEPS/NMATS, " + @title_year + " " + @start_month_name + "-" + @month_name], :style => @tab_subheader
      sheet.add_row []
      sheet.add_row []
      sheet.add_row ["","","Selected Customer*","='GEPS Volume & Declines'!$C$4","Selected Product*","='GEPS Volume & Declines'!$E$4", "", "*Alter these choices by using the dropdowns on the 'GEPS Volume & Declines' tab"], :style => @wrap_text
      sheet.add_row []
      
      sheet.merge_cells "H6:O6"
      
      @header = ["", "", "Rate Groups"]
      @header_style = []
      @header2_style = []
      @rategroup_row_style = []
      @last_rategroup_row_style = []
      (@geps_rategroups.length-2).times do
        @header << ""
        @header_style << @row_header_merge
        @header2_style << @row_header_middle
        @rategroup_row_style << @row_styling_vol
        @last_rategroup_row_style << @last_row_vol
      end
      
      @header_style << @row_header_merge
      @header2_style << [@row_header_middle, @row_header_middle]
      @rategroup_row_style << [@row_styling_vol, @row_styling_vol]
      @last_rategroup_row_style << [@last_row_vol, @last_row_vol]
      
      sheet.add_row [@header].flatten, :style => [nil, nil, @header_style].flatten
      @geps_rategroups.shift
      sheet.add_row ["", "Weight Step", @geps_rategroups, "All"].flatten, :style => [nil, @row_header_middle, @header2_style].flatten
      
      
      @sum_letter_arr = ('F'..'Z').to_a
      @sum = ""
      (@date_arr.length).times do |count|
        @sum_letter = @sum_letter_arr[count]
        if count == 0
          @sum = "('PQW Report Data'!$F$4:$F$11233)"
        elsif count == @date_arr.length-1
          @sum += "+('PQW Report Data'!$#{@sum_letter}$4:$#{@sum_letter}$11233))"
        else
          @sum += "+('PQW Report Data'!$#{@sum_letter}$4:$#{@sum_letter}$11233)"
        end
      end
      
      
      
      
      @weightsteps.each_with_index do |w, index|
        if w != 'All'
          @row_arr = ["", w]
          @letter_arr = ('C'..'Z').to_a
          @letter_arr = @letter_arr.push('AA', 'AB')
          
          @geps_rategroups.each do |g|
            @letter = @letter_arr.shift()
            @temp = %Q|=IF(AND($D$6="All",$F$6="All"),SUMPRODUCT(('PQW Report Data'!$C$4:$C$11233=#{@letter}$9)*('PQW Report Data'!$E$4:$E$11233=$B#{index+9})*(#{@sum}),
                    IF($D$6="All",SUMPRODUCT(('PQW Report Data'!$D$4:$D$11233='GEPS Volume & Declines'!$E$4)*('PQW Report Data'!$C$4:$C$11233=#{@letter}$9)*('PQW Report Data'!$E$4:$E$11233=$B#{index+9})*(#{@sum}),
                    IF($F$6="All",SUMPRODUCT(('PQW Report Data'!$B$4:$B$11233='GEPS Volume & Declines'!$C$4)*('PQW Report Data'!$C$4:$C$11233=#{@letter}$9)*('PQW Report Data'!$E$4:$E$11233=$B#{index+9})*(#{@sum}),
                    SUMPRODUCT(('PQW Report Data'!$B$4:$B$11233='GEPS Volume & Declines'!$C$4)*('PQW Report Data'!$D$4:$D$11233='GEPS Volume & Declines'!$E$4)*('PQW Report Data'!$C$4:$C$11233=#{@letter}$9)*('PQW Report Data'!$E$4:$E$11233=$B#{index+9})*(#{@sum}))))|
            @row_arr << @temp
          end
            @row_arr << ["=SUM(C#{index+9}:AA#{index+9})"]
          if index == (@weightsteps.length-1)
            sheet.add_row [@row_arr].flatten, :style => [nil, @last_row_middle, @last_rategroup_row_style].flatten
          else
            sheet.add_row [@row_arr].flatten, :style => [nil, @row_styling_middle, @rategroup_row_style].flatten  
          end
        end
      end
      
      @letter_arr = ('C'..'Z').to_a
      @letter_arr = @letter_arr.push('AA')
      @totals = ["", "Total"]
      
      @letter_arr.each do |l|
        @temp = "=SUM(#{l}10:#{l}80)"
        @totals << @temp
      end
      
      sheet.add_row [@totals].flatten, :style => [nil, @row_styling_middle, @rategroup_row_style].flatten
      
      sheet.merge_cells "C8:AA8"
      sheet.column_info[8].hidden = true
      sheet.column_info[9].hidden = true
      sheet.column_info[10].hidden = true
      
      sheet.column_widths 2, 10, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11
      
    end
    
    ###VOLUME CHANGE BY WS&RG###
    workbook.add_worksheet(name: "Monthly Volume Change WS & RG") do |sheet|
      
      ###Hide gridlines on page
      sheet.sheet_view.show_grid_lines= false
      sheet.sheet_pr.tab_color = @tab3
      
      sheet.add_row []
      sheet.add_row ["","Rate Group and Weight Step Month to Month Changes Report"], :style => @tab_header
      sheet.add_row ["",  "Source: GEPS/NMATS, " + @title_year + " " + @start_month_name + "-" + @month_name], :style => @tab_subheader
      sheet.add_row []
      sheet.add_row []
      sheet.add_row ["","","Selected Customer*","='GEPS Volume & Declines'!$C$4","Selected Product*","='GEPS Volume & Declines'!$E$4", "", "*Alter these choices by using the dropdowns on the 'GEPS Volume & Declines' tab"], :style => @wrap_text
      sheet.add_row []
      
      sheet.merge_cells "H6:O6"
      
      @header = ["", "", "Rate Groups"]
      @header_style = []
      @header2_style = []
      @rategroup_row_style = []
      @last_rategroup_row_style = []
      (@geps_rategroups.length-1).times do
        @header << ""
        @header_style << @row_header_merge
        @header2_style << @row_header_middle
        @rategroup_row_style << @row_styling_vol
        @last_rategroup_row_style << @last_row_vol
      end
      
      @header_style << @row_header_merge
      @header2_style << [@row_header_middle, @row_header_middle]
      @rategroup_row_style << [@row_styling_vol, @row_styling_vol]
      @last_rategroup_row_style << [@last_row_vol, @last_row_vol]
      
      sheet.add_row [@header].flatten, :style => [nil, nil, @header_style].flatten
      
      sheet.add_row ["", "Weight Step", @geps_rategroups, "All"].flatten, :style => [nil, @row_header_middle, @header2_style].flatten
      
      
      
      @weightsteps.each_with_index do |w, index|
        if w != 'All'
          @row_arr = ["", w]
          @letter_arr = ('C'..'Z').to_a
          @letter_arr = @letter_arr.push('AA', 'AB')
          
          @geps_rategroups.each do |g|
            @letter = @letter_arr.shift()
            @temp = %Q|=IF(AND($D$6="All",$F$6="All"),SUMPRODUCT(('PQW Report Data'!$C$4:$C$11233=#{@letter}$9)*('PQW Report Data'!$E$4:$E$11233=$B#{index+9})*(('PQW Report Data'!K$4:K$11233)-('PQW Report Data'!J$4:J$11233))),
                    IF($D$6="All",SUMPRODUCT(('PQW Report Data'!$D$4:$D$11233='GEPS Volume & Declines'!$E$4)*('PQW Report Data'!$C$4:$C$11233=#{@letter}$9)*('PQW Report Data'!$E$4:$E$11233=$B#{index+9})*(('PQW Report Data'!K$4:K$11233)-('PQW Report Data'!J$4:J$11233))),
                    IF($F$6="All",SUMPRODUCT(('PQW Report Data'!$B$4:$B$11233='GEPS Volume & Declines'!$C$4)*('PQW Report Data'!$C$4:$C$11233=#{@letter}$9)*('PQW Report Data'!$E$4:$E$11233=$B#{index+9})*(('PQW Report Data'!K$4:K$11233)-('PQW Report Data'!J$4:J$11233))),
                    SUMPRODUCT(('PQW Report Data'!$B$4:$B$11233='GEPS Volume & Declines'!$C$4)*('PQW Report Data'!$D$4:$D$11233='GEPS Volume & Declines'!$E$4)*('PQW Report Data'!$C$4:$C$11233=#{@letter}$9)*('PQW Report Data'!$E$4:$E$11233=$B#{index+9})*(('PQW Report Data'!K$4:K$11233)-('PQW Report Data'!J$4:J$11233))))))|
            @row_arr << @temp
          end
            @row_arr << "=SUM(C#{index+9}:AA#{index+9})"
            
          if index == (@weightsteps.length-1)
            sheet.add_row [@row_arr].flatten, :style => [nil, @last_row_middle, @last_rategroup_row_style].flatten
          else
            sheet.add_row [@row_arr].flatten, :style => [nil, @row_styling_middle, @rategroup_row_style].flatten
          end
        end
      end
      
      @letter_arr = ('C'..'Z').to_a
      @letter_arr = @letter_arr.push('AA')
      @totals = ["", "Total"]
      @letter_arr.each do |l|
        @temp = "=SUM(#{l}10:#{l}80)"
        @totals << @temp
      end
      sheet.add_row [@totals].flatten, :style => [nil, @row_styling_middle, @rategroup_row_style].flatten
      
      # Apply conditional formatting in the worksheet
      sheet.add_conditional_formatting("C10:AB81", { :type => :cellIs,
                                          :operator => :lessThan,
                                          :formula => '0',
                                          :dxfId => unprofitable,
                                          :priority => 1 })
      sheet.add_conditional_formatting("C10:AB81", { :type => :cellIs,
                                          :operator => :greaterThan,
                                          :formula => '0',
                                          :dxfId => profitable,
                                          :priority => 1 })
      
      sheet.column_info[8].hidden = true
      sheet.column_info[9].hidden = true
      sheet.column_info[10].hidden = true
      
      sheet.merge_cells "C8:AA8"
      
      sheet.column_widths 2, 10, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11
    end
     
    ###DESTINATIONS MONTH OVER MONTH###
    workbook.add_worksheet(name: "Top Destinations by Monthly Vol") do |sheet|
      
      ###Hide gridlines on page
      sheet.sheet_view.show_grid_lines= false
      sheet.sheet_pr.tab_color = @tab3
      
      sheet.add_row []
      sheet.add_row ["","Top Destinations by Monthly Volume"], :style => @tab_header
      sheet.add_row ["", "Source: GEPS/NMATS, "+ @title_year + " " + @start_month_name + "-" + @month_name], :style => @tab_subheader
      sheet.add_row []
      sheet.add_row []
      
       @header = ["", "", "Calendar Month"]
      (@date_arr.length-1).times do
        @temp = ""
        @header << @temp
      end
      sheet.add_row [@header, "Prior Month Change", ""].flatten, :style => [nil, @row_header, @style_merge, @row_header_merge, @row_header_merge].flatten
      sheet.add_row ["", "Destination", @date_arr, "Gross Change", "Percent Change"].flatten, :style => [nil, @row_header, @date_style, @row_header_middle, @row_header_middle].flatten
      
      
      15.times do |count|
        @letter_arr = ('C'..'Z').to_a
        @row_arr = []
         (@date_arr.length).times do
          @letter = @letter_arr.shift()
          @temp = %Q|=VLOOKUP($B#{count+8}, Destinations_Ranked!$A$4:$AI$221, MATCH("_"&'Top Destinations by Monthly Vol'!#{@letter}$7, Destinations_Ranked!$A$3:$AO$3, 0), 0)|
          @row_arr << @temp
         end
         
        @letter_arr = ('C'..'Z').to_a
        sheet.add_row [count+1, %Q|=Destinations_Ranked!A#{count+4}|, @row_arr,
                       %Q|=#{@letter_arr[@date_arr.length-1]}#{count+8}-#{@letter_arr[@date_arr.length-2]}#{count+8}|,
                       %Q|=IFERROR(#{@letter_arr[@date_arr.length]}#{count+8}/#{@letter_arr[@date_arr.length-2]}#{count+8},0)|].flatten,
                      :style => [@row_num, @row_styling_middle, @vol_style, @row_styling_vol, @row_styling_per].flatten
        
      end
      
      @col_arr = ('B'..'Z').to_a
      @last = @col_arr[@date_arr.length]
      @next_to_last = @col_arr[@date_arr.length-1]
      sheet.add_chart(Axlsx::Bar3DChart, :start_at => "L7", :end_at => "R20", :title=> "PQW Volume for Top 15 Top Destinations by Monthly Vol", :barDir => :col) do |bars|
              bars.add_series :data => sheet["#{@next_to_last}8:#{@next_to_last}22"], :labels => sheet["B8:B22"], :title => sheet["#{@next_to_last}7"]
              bars.add_series :data => sheet["#{@last}8:#{@last}22"], :labels => sheet["B8:B22"], :title => sheet["#{@last}7"]
              bars.catAxis.label_rotation = -43
              bars.valAxis.gridlines = false
              bars.catAxis.gridlines = false
            end
      
      @form_arr = ('C'..'Z').to_a
      # Apply conditional formatting in the worksheet
      sheet.add_conditional_formatting("#{@form_arr[@date_arr.length]}9:#{@form_arr[@date_arr.length+1]}37", { :type => :cellIs,
                                          :operator => :lessThan,
                                          :formula => '0',
                                          :dxfId => unprofitable,
                                          :priority => 1 })
      sheet.add_conditional_formatting("#{@form_arr[@date_arr.length]}9:#{@form_arr[@date_arr.length+1]}37", { :type => :cellIs,
                                          :operator => :greaterThan,
                                          :formula => '0',
                                          :dxfId => profitable,
                                          :priority => 1 })
      
      sheet.merge_cells "C6:#{@last}6"
      sheet.merge_cells "#{@col_arr[@date_arr.length+1]}6:#{@col_arr[@date_arr.length+2]}6"
      
      sheet.column_widths 2, 14, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10
    end
    
    ###CUSTOMERS MONTH OVER MONTH GEPS###
    workbook.add_worksheet(name: "Top Customers by Monthly Vol") do |sheet|
      
      ###Hide gridlines on page
      sheet.sheet_view.show_grid_lines= false
      sheet.sheet_pr.tab_color = @tab3
      
      sheet.add_row []
      sheet.add_row ["","Top Customers by Monthly Volume"], :style => @tab_header
      sheet.add_row ["", "Source: GEPS/NMATS, "+ @title_year + " " + @start_month_name + "-" + @month_name], :style => @tab_subheader
      sheet.add_row []
      sheet.add_row []
      
       @header = ["", "", "Calendar Month"]
      (@date_arr.length-1).times do
        @temp = ""
        @header << @temp
      end
      sheet.add_row [@header, "Prior Month Change", ""].flatten, :style => [nil, @row_header, @style_merge, @row_header_merge, @row_header_merge].flatten
      sheet.add_row ["", "Destination", @date_arr, "Gross Change", "Percent Change"].flatten, :style => [nil, @row_header, @date_style, @row_header_middle, @row_header_middle].flatten
      
      
      (@geps_pqws.length-1).times do |count|
        @letter_arr = ('C'..'Z').to_a
        @row_arr = []
         (@date_arr.length).times do
          @letter = @letter_arr.shift()
          @temp = %Q|=IFERROR(VLOOKUP($B#{count+8}, Customers_Ranked!$B$4:$AQ$1920, MATCH("_"&'Top Customers by Monthly Vol'!#{@letter}$7, Customers_Ranked!$B$3:$AR$3, 0), 0),"")|
          @row_arr << @temp
        end
        
        @letter_arr = ('C'..'Z').to_a
        @letter = @letter_arr[count]
        sheet.add_row [count+1, %Q|=Customers_Ranked!B#{count+4}|, @row_arr,
                       %Q|=#{@letter_arr[@date_arr.length-1]}#{count+8}-#{@letter_arr[@date_arr.length-2]}#{count+8}|,
                       %Q|=IFERROR(#{@letter_arr[@date_arr.length]}#{count+8}/#{@letter_arr[@date_arr.length-2]}#{count+8},0)|].flatten,
                       :style => [@row_num, @row_styling_middle, @vol_style, @row_styling_vol, @row_styling_per].flatten
      end
      
      @col_arr = ('B'..'Z').to_a
      @last = @col_arr[@date_arr.length]
      @next_to_last = @col_arr[@date_arr.length-1]
      sheet.add_chart(Axlsx::Bar3DChart, :start_at => "L8", :end_at => "R18", :title=> "GEPS Volume for Top 5 PQWs Month over Month", :barDir => :col) do |bars|
              bars.add_series :data => sheet["#{@next_to_last}8:#{@next_to_last}12"], :labels => sheet["B8:B12"], :title => sheet["#{@next_to_last}7"]
              bars.add_series :data => sheet["#{@last}8:#{@last}12"], :labels => sheet["B8:B12"], :title => sheet["#{@last}7"]
              bars.catAxis.label_rotation = 15
              bars.valAxis.gridlines = false
              bars.catAxis.gridlines = false
      
      
      sheet.merge_cells "C6:#{@last}6"
      sheet.merge_cells "#{@col_arr[@date_arr.length+1]}6:#{@col_arr[@date_arr.length+2]}6"
      
      @form_arr = ('C'..'Z').to_a
      # Apply conditional formatting in the worksheet
      sheet.add_conditional_formatting("#{@form_arr[@date_arr.length]}8:#{@form_arr[@date_arr.length+1]}19", { :type => :cellIs,
                                          :operator => :lessThan,
                                          :formula => '0',
                                          :dxfId => unprofitable,
                                          :priority => 1 })
      sheet.add_conditional_formatting("#{@form_arr[@date_arr.length]}8:#{@form_arr[@date_arr.length+1]}19", { :type => :cellIs,
                                          :operator => :greaterThan,
                                          :formula => '0',
                                          :dxfId => profitable,
                                          :priority => 1 })
      end
      
      sheet.column_widths 2, 14, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10
    end
    
    ###Rategroup Breakouts###
    workbook.add_worksheet(name: "Rate Group Country Reference") do |sheet|
      
      ###Hide gridlines on page
      sheet.sheet_view.show_grid_lines= false
      
      sheet.add_row []
      sheet.add_row ["","Rate Group Country Reference"], :style => @tab_header
      sheet.add_row ["","By Product Group"], :style => @tab_subheader
      sheet.add_row []
      sheet.add_row ["", "", "Groupings for IPA/ISAL/CEP", "", "", "Groups for PMEI/PMI", "", "", "Groupings for FCPIS"], :style => [nil,@row_header,@row_header,nil,@row_header,@row_header,nil,@row_header,@row_header]
      sheet.add_row ["", "Rategroup", "Countries","", "Rategroup", "Countries", "", "Rategroup", "Countries"], :style => [nil, @row_header_middle, @row_header_middle, nil, @row_header_middle, @row_header_middle, nil, @row_header_middle, @row_header_middle]
      
      @ipa_isal_cep = [      
      "Canada",
      "Mexico",
      "Great Britain & Northern Ireland",
      "Germany",
      "France",
      "Japan",
      "Italy",
      "Spain",
      "Australia, New Zealand",
      "Argentina, Brazil",
      "Hong Kong, South Korea, Singapore",
      "Austria, Belgium, Denmark, Finland, Netherlands, Norway, Poland, Sweden, Switzerland",
      "Greece, Ireland, Israel, Portugal, Slovenia",
      "China, India, Phillipines, South Africa, Taiwan, Thailand",
      "Greenland, Andorra, Iceland, Luxembourg, Gibraltar, Liechtenstein",
      "Albania, Belarus, Bosnia, Bulgaria, Croatia, Czech Republic, Estonia, Hungary, Kosovo, Latvia, Lithuania, Romania, Russia, Serbia, Turkey",
      "Vietnam, Fiji, Nepal, Mongolia, Malaysia, Laos, Indonesia, Western Africa",
      "Venezuela, Peru, and the Caribbean",
      "Middle East & Eastern Africa, as well as Ukraine"
      ]
      
      @pmi_pmei = [
        "Canada",
        "Mexico",
        "Hong Kong & South Korea",
        "Albania, Armenia, Azerbaijan, Belarus, Bosnia and Herzegowina, Bulgaria, Croatia, Cyprus, Czech Republic, Estonia, Georgia, Republic of, Hungary, Latvia, Lithuania, Macedonia, Republic of Moldova, Poland, Romania, Russia, Turkey, Ukraine", 
        "Andorra, Austria, Belgium, Denmark, Faroe Islands, Finland, Greece, Iceland, Ireland, Italy, Liechtenstein, Luxembourg, Malta, Norway, Portugal, Republic of Serbia, San Marino, Republic of Slovak Republic (Slovakia), Slovenia, Spain, Sweden, Switzerland, Vatican City",
        "Bangladesh, Bhutan, Brunei Darussalam, Burma, Cambodia, Fiji, French Polynesia, India, Indonesia, Kazakhstan, Kiribati, Kyrgyzstan, Laos, Macao, Malaysia, Maldives, Mongolia, Nauru, Nepal, New Caledonia, Pakistan, Papua New Guinea, Philippines, Samoa, Singapore, Solomon Islands, Sri Lanka, Taiwan, Tajikistan, Thailand, Tonga, Turkmenistan, Uzbekistan, Vanuatu,, Vietnam",
        "Angola, Benin, Botswana, Burkina Faso, Burundi, Cameroon, Cape Verde, Central African Republic, Chad, Congo, Democratic Republic of the Congo, Republic of the Cote d'Ivoire (Ivory Coast), Djibouti, Equatorial Guinea, Eritrea, Gabon, Ghana, Guinea, Guinea-Bissau, Kenya, Lesotho, Liberia, Madagascar, Malawi, Mali, Mauritania, Mauritius, Mozambique, Namibia, Niger, Nigeria, Rwanda, Sao Tome and Principe, Senegal, Seychelles, Sierra Leone, South Africa, Sudan, Swaziland, Tanzania, Togo, Uganda, Zambia, Zimbabwe",
        "Algeria, Bahrain, Egypt, Ethiopia, Iraq, Israel, Jordan, Kuwait, Lebanon, Morocco, Oman, Qatar, Saudi Arabia, Syrian Arab Republic (Syria), Tunisia, United Arab Emirates, Yemen",
        "Anguilla, Argentina, Aruba, Bahamas, Barbados, Belize, Bermuda, Bolivia, Bonaire, Sint Eustatius, and Saba, Cayman Islands, Chile, Colombia, Costa Rica, Curacao, Dominica, Dominican Republic, Ecuador, El Salvador, French Guiana, Grenada, Guadeloupe, Guatemala, Guyana, Haiti, Honduras, Jamaica, Martinique, Netherlands Antilles, Nicaragua, Panama, Paraguay, Peru, Saint Kitts and Nevis, Saint Lucia, Saint Vincent and the Grenadines, Sint Maarten, Trinidad and Tobago, Turks and Caicos Islands, Uruguay, Venezuela",        "",
        "Australia, Christmas Island, New Zealand",
        "Great Britain and Northern Ireland",
        "Japan",
        "France",
        "China",
        "Brazil",
        "Germany",
        "Netherlands",
        ""
        ]
      
      @fcpis = [
        "Canada",
        "Mexico",
        "Australia, China, Christmas Island, Hong Kong, Japan, Korea, Republic of (South Korea)",
        "Albania, Armenia, Azerbaijan, Belarus, Bosnia and Herzegowina, Bulgaria, Croatia, Cyprus, Czech Republic, Estonia, Georgia, Republic of Hungary, Latvia, Lithuania, Macedonia, Republic of Moldova, Poland, Romania, Russia, Saint Pierre and Miquelon, Turkey, Ukraine",
        "Andorra, Austria, Belgium, Denmark, Faroe Islands, Finland, France, Germany, Gibraltar, Great Britain and Northern Ireland, Greece, Greenland, Iceland, Ireland, Israel, Italy, Kosovo, Kosovo, Republic of, Liechtenstein, Luxembourg, Malta, Montenegro, Netherlands, Norway, Portugal, Republic of Serbia, San Marino, Serbia, Republic of, Serbia, Republic of, Serbia, Republic of, Slovak Republic (Slovakia), Slovenia, Spain, Sweden, Switzerland, Vatican City",
        "Afghanistan, Bangladesh, Bhutan, Brunei Darussalam, Burma, Cambodia, Fiji, French Polynesia, India, Indonesia, Kazakhstan, Kiribati, Korea, Democratic People's Republic of (North Korea), Kyrgyzstan, Laos, Macao, Malaysia, Maldives, Mongolia, Nauru, Nepal, New Caledonia, New Zealand, Pakistan, Papua New Guinea, Philippines, Pitcairn Island, Samoa, Singapore, Solomon Islands, Sri Lanka, Taiwan, Tajikistan, Thailand, Timor-Leste, Democratic Republic of, Tonga, Turkmenistan, Tuvalu, Uzbekistan, Vanuatu, Vietnam, Wallis and Futuna Islands",
        "Afghanistan, Angola, Ascension, Benin, Botswana, Burkina Faso, Burundi, Cameroon, Cape Verde, Central African Republic, Chad, Comoros, Congo, Democratic Republic of the, Congo, Republic of the, Cote d'Ivoire (Ivory Coast), Djibouti, Equatorial Guinea, Eritrea, Gabon, Gambia, Ghana, Guinea, Guinea-Bissau, Kenya, Lesotho, Liberia, Madagascar, Malawi, Mali, Mauritania, Mauritius, Mozambique, Namibia, Niger, Nigeria, Rwanda, Saint Helena, Sao Tome and Principe, Senegal, Seychelles, Sierra Leone, South Africa, Sudan, Swaziland, Tanzania, Togo, Tristan da Cunha, Uganda, Zambia, Zimbabwe",
        "Algeria, Bahrain, Egypt, Ethiopia, Iran, Iraq, Jordan, Kuwait, Lebanon, Libya, Morocco, Oman, Qatar, audi Arabia, Syrian Arab Republic (Syria), Tunisia, United Arab Emirates, Yemen",        
        "Anguilla, Antigua and Barbuda, Argentina, Aruba, Bahamas, Barbados, Belize, Bermuda, Bolivia, Bonaire, Sint Eustatius, and Saba, Brazil, British Virgin Islands, Cayman Islands, Chile, Colombia, Costa Rica, Cuba, Curacao, Dominica, Dominican Republic, Ecuador, El Salvador, Falkland Islands, French Guiana, Grenada, Guadeloupe, Guatemala, Guyana, Haiti, Honduras, Jamaica, Martinique, Montserrat, Netherlands Antilles, Nicaragua, Panama, Paraguay, Peru, Reunion, Saint Kitts and Nevis, Saint Lucia, aint Vincent and the Grenadines, Sint Maarten, Suriname, Trinidad and Tobago, Turks and Caicos Islands, Uruguay, Venezuela",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        ""
        ]
      
      @ipa_isal_cep.each_with_index do |country,index|
        @row_arr = []
        @temp = ["", index+1, country, ""]
        if index <= @pmi_pmei.length-1
          if index <= @fcpis.length-1
          @temp2 = [index+1, @pmi_pmei[index], "", index+1, @fcpis[index], ""]
          else
          @temp2 = [index+1, @pmi_pmei[index]]
          end
          sheet.add_row [@temp, @temp2].flatten, :style => @plain_row
        else
          sheet.add_row [@temp].flatten, :style => @plain_row
        end
      end
    
      sheet.column_widths 5, 10, 40, 5, 10, 80, 5, 10, 80  
    end
    
    
    ###DATA>>>###
    workbook.add_worksheet(name: "Data >>>>") do |sheet|
      
    sheet.sheet_pr.tab_color = @tab_blu

      sheet.add_row []
      sheet.add_row []
      sheet.add_row ["Data >>>>"]
    end
    
    ###PQW REPORT DATA###
    workbook.add_worksheet(name: "PQW Report Data") do |sheet|
      sheet.add_row ["", "The SaS System", "", "Source: Geps_Zero_pqwcombined"]
      
    end
    
    ###POSTAL ONE DATA###
    workbook.add_worksheet(name: "Postal One PQW Report") do |sheet|
      sheet.add_row ["", "The SaS System", "", "Source: Postalone_pqw_zerocombined","","IMPORANT, CHANGE RATE GROUP 0 TO WORLDWIDE"]
      
    end
    
    ###CUSTOMERS RANKED###
    workbook.add_worksheet(name: "Customers_Ranked") do |sheet|
      sheet.add_row ["", "The SaS System", "", "source: ranked_sort_customer_null"]
      sheet.add_row []
      sheet.add_row []
      
    end
    
    ###DESTINATIONS RANKED###
    workbook.add_worksheet(name: "Destinations_Ranked") do |sheet|
      sheet.add_row ["", "The SaS System", "", "source: ranked_sort_country_null"]
      sheet.add_row []
      sheet.add_row []
      
    end
    
    ###TOP GRANULAR LEVEL DECLINES###
    workbook.add_worksheet(name: "Top Granular Level Declines") do |sheet|
      sheet.add_row ["", "The SaS System", "", "Source: WSRG_Declines_Ranked"]
    end
    
    package.serialize("pqw_report.xlsx")
  end
end
