require 'zip/zip'
require 'zip/zipfilesystem'
require 'net/http'
require 'uri'
require 'rubygems'
require 'spreadsheet'

#url = URI.parse('http://www.microtron.zp.ua')
#res = Net::HTTP.start(url.host, url.port) {|http|
 # http.get('/price/mt1.zip')
#}
#f=File.new("1.zip",'wb')
#f.write(res.body)
#if File.exists?('1.xls')
 # FileUtils.remove_file("1.xls")
#end
#Zip::ZipFile::open("1.zip") {
 # |zf|
  #zf.each { |e|
   # path = File.join('1.xls')
    #FileUtils.mkdir_p(File.dirname(path))
    #zf.extract(e, path)
  #}
#}
#FileUtils.remove_file("1.zip")

excel = Spreadsheet.open('1.xls').worksheet 0
groups=0
brand=1
item=2
price=3
subgroup=''
yamlfile = File.open('1.yml', 'w')
  excel.each do |rows|
    if rows.default_format and rows[groups].to_i==0
        yamlfile << rows[item].to_s << "\n" if rows.default_format.pattern_fg_color==:border and rows.idx>1
     
    end 
	 if rows.default_format and rows[groups].to_i==0
				  yamlfile << rows[brand].to_s << "\n" if rows.default_format.pattern_fg_color==:border and rows.idx>1
			  
	 end
	 if rows.default_format and rows[groups].to_i==0
		     yamlfile << rows[price].to_s << "\n" if rows.default_format.pattern_fg_color==:border and rows.idx>1

		  
	 end 
  
		 if rows.default_format and rows.default_format.pattern_fg_color==:border and rows[groups].to_i!=0
		   yamlfile << rows[groups] << "  " << rows[item].to_s << "\n" << "  " << rows[price]<< "\n" << "  " << rows[brand]<< "\n" if subgroup==''
		   yamlfile <<  rows[groups] << "    " << rows[item].to_s << "\n" << "    " << rows[price]<< "\n" << " "<< rows[brand]<< "\n" unless subgroup==''
		 end
   
  end
#  FileUtils.remove_file("1.xls")
