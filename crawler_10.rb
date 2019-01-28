require 'rubygems'
require 'nokogiri'
require 'open-uri'
require 'spreadsheet'
require 'rufus-scheduler'
require 'net/http'

$No_of_cur = 9
cur_price = Array.new

def open(url)
  Net::HTTP.get(URI.parse(url))
end


def get_price(cur_price_in)

	hren = Array.new
	hrenn=Array.new
	new_hren = Array.new
	page_content = open('https://min-api.cryptocompare.com/data/pricemulti?fsyms=ETH,BTC,NEO,BNB,XRP,ADA,QTUM,LTC,BCH&tsyms=USD')
	hren = page_content.split(",")
	#hren = page_content.split(",")
#puts 7.chr

	for i in 0..$No_of_cur-1 do 
		#puts "i=#{i}"
		cur_price_in[i] = Array.new
		hrenn=hren[i].split(":")
		if i == 0 
			new_hren[0] = hrenn[0][2..-2]
			new_hren[2] = hrenn[2][0..-2]
		else 
			if i == $No_of_cur
				new_hren[0] = hrenn[0][1..-2]
				new_hren[2] = hrenn[2][0..-3]
			else
				new_hren[0] = hrenn[0][1..-2]
				new_hren[2] = hrenn[2][0..-2]
			end
		end
	#puts "new hren=#{new_hren}"
    	cur_price_in[i][0] = new_hren[0]
    	cur_price_in[i][2] = new_hren[2]
    	#puts "new hren2=#{cur_price_in[i]}"
	end
	return cur_price_in
end

def write_price(cur_price_in)
	book = Spreadsheet.open('stat_template.xls')
	for i in 0..$No_of_cur-1 do
	 	sheet1 = book.worksheet(cur_price_in[i][0].downcase)
	 	sheet1.each do |row|
   
    		if row[0] !=nil
    			# puts row[2]
    			next
    		else
    			row[0]=Time.now.to_s[0..-7]
    			row[1]=cur_price_in[i][2].to_f
    			break
    		end
    		#puts "row[2]=#{row[2]}"
    	end
    end

	File.delete('stat_template.xls')
	book.write('stat_template.xls')
	return cur_price_in
end

	#get_price(cur_price)
	#write_price(cur_price)


 #main program
#=begin
scheduler = Rufus::Scheduler.new
scheduler.every '5m' do
	puts Time.now

	get_price(cur_price)
	write_price(cur_price)

end
scheduler.join
#=end
