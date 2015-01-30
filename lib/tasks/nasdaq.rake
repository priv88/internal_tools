namespace :nasdaq_monthly do 
  desc "scrap the info from nasdaq on a monthly basis"
  task :scraper => :environment do 
    require 'rubygems'
    require 'nokogiri'
    require 'pp'
    require 'open-uri'
    require 'writeexcel'
    require 'mechanize'
    require 'pry-byebug'

    class Float
      def even?
      self%1==0 && self.to_i.even?
      end
    end

    class String
      def string_between_markers marker1, marker2
        self[/#{Regexp.escape(marker1)}(.*?)#{Regexp.escape(marker2)}/m, 1]
      end
    end

    def find_prior_month(time)
      month = time.strftime('%m').to_i
      
      if month == 1 
        prior_month = 12
      else
        prior_month = month - 1
      end
      
      return prior_month.to_s
    end

    def find_year(time)
      year = time.strftime("%Y").to_i
      month = time.strftime("%m").to_i
      puts month
      if month == 1
        year = year - 1
      end
      return year.to_s
    end 


    if Time.now.strftime('%d') == "01"

    # month = [Time.now.strftime('%m')]
      month = [prior_month(Time.now)]
      year = [find_year(Time.now)]
      
      # Reason for crawling list of recently filed IPOs: these are the companies with fresh, public financials; the finer categories "upcoming", "latest", and "recently withdrawn" can be identified within the company pages - this can be achieved by periodically refreshing the profile pages of <companies that have filed for IPOs>

      agent = Mechanize.new
      agent.user_agent_alias = 'Windows Chrome'
      agent.ssl_version = "SSLv3"
      agent.verify_mode = OpenSSL::SSL::VERIFY_NONE

      # .genTable > table:nth-child(1)

      link1 = Array.new

      year.each do |year|
        puts year
        month.each do |month|
          puts month
          # binding.pry
          page = Nokogiri::HTML(open("http://www.nasdaq.com/markets/ipos/activity.aspx?tab=filings&month=#{year}-#{month}")).css('table')[2].css('tr').css('a') 
          # page =  page.css('table')[2].css('tr').css('a')
          # links = ÷ 
          page.xpath("//a/@href").map(&:to_s).uniq.each do |link|
            if link.to_s.match("http://www.nasdaq.com/markets/ipos/company")
              puts link.to_s
              link1 << link.to_s
            else
              next
            end
          end
        end
      end

      workbook = WriteExcel.new('Nasdaq_Filings.xls')
      worksheet1 = workbook.add_worksheet

      headers = ["Company Name", "Company Address", "Company Phone", "Company Website", "CEO", "Latest Employees", "State of Inc", "Fiscal Year End", "Status", "Proposed Symbol","Exchange","Share Price", "Shares Offered", "Offer Amount", "Total Expense", "Shares Over Alloted", "Shareholders Shares Offered", "Shares Outstanding", "Lockup Period (days", "Lockup Expiration", "Quiet Period Expiration", "CIK"]
      worksheet2 = workbook.add_worksheet
      headers.each_index do |index|
        header = headers[index]
        worksheet2.write(0,index,header)
      end

      summary_count = 1
      link1.each do |url|
        doc = Nokogiri::HTML(open(url))
        company_name = doc.css("table")[2].css("td")[1].text
        puts company_name
        (0..21).each do |index|
          worksheet2.write(summary_count,index,doc.css("table")[2].css("td")[index * 2 + 1].text)
        end
        summary_count += 1
      end



      count_1 = 0 

      link1.each do |url|
        doc = Nokogiri::HTML(open(url))
        worksheet = workbook.add_worksheet
        company_name = doc.css("table")[2].css("td")[1].text
        worksheet1.write(count_1,0,"#{company_name}")

        (0..21).each do |index|
          worksheet.write(index,0,doc.css("table")[2].css("td")[index * 2].text)
          worksheet.write(index,1,doc.css("table")[2].css("td")[index * 2 + 1].text)
        end

        worksheet.write(23,0,"Company Description")
        worksheet.write(23,1,doc.css(".ipo-comp-description").css("pre").text + doc.css('div#read_more_div_toggle1').css('pre').text)
        worksheet.write(25,0,"Use of Proceeds")
        worksheet.write(25,1,doc.css('div#infoTable_2').css('pre').text)


        link2 = Array.new
        doc.css('table')[5].css('tr').css('a').xpath("//a/@href").map(&:to_s).uniq.each do |link|
          if link.to_s.include? "/markets/ipos/filing.ashx?filingid="
            link2 << link.to_s
          else
            next
          end
        end

        count=0
        row=27
        doc.css('table')[5].css('tr').children.each do |item|
          item = item.text.delete("\n").delete("\r").delete("\t").strip
          if item.empty? == false
            if count % 4 == 0
              worksheet.write(row,0,item)
            elsif count % 4 == 1
              worksheet.write(row,1,item)
            elsif count % 4 == 2
              worksheet.write(row,2,item)
            else
              worksheet.write(row,3,item)
              row+=1
            end
            count+=1
          else
            next
          end
        end

        dummyrow=0
        (1...count/4).each do|index|
          worksheet.write(28+dummyrow,4,"http://www.nasdaq.com"+link2[index-1])
          dummyrow+=1
        end
        
        expert = Array.new
        doc.css('table')[6].css('td').children.each do |item|
          expert << item.to_s
        end

        dummycount=0
        dummycount1=0
        expert.each do |item|
          if dummycount.even?
            worksheet.write(28+dummyrow+dummycount1+1,0,item)
          else
            worksheet.write(28+dummyrow+dummycount1+1,1,item.string_between_markers(">","<"))
            dummycount1+=1
          end
          dummycount+=1
        end

        worksheet1.write_url(count_1,1,"internal:Sheet#{count_1+3}!A1")
        count_1 +=1
        puts "@"

      end
      workbook.close
      # binding.pry
      type = "filings"
      TaskMailer.send_nasdaq_email(type).deliver!
      puts "filings"
    end
  end
end