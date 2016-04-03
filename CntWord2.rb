# coding: UTF-8
Encoding.default_internal = "utf-8"
# Encoding.default_external = "gbk"

require 'yaml'
require 'fileutils'
require 'pp'
require_relative 'word2'
require_relative 'excel'

class SummaryParagraph
	attr :all
	def initialize detect
		@detect = detect
		@all = []
		Dir["**/*.doc*"].each{|one|
			# pp one
			if /\//=~ one #只找子路径下的Word文件
				# pp one
				file_path="#{Dir.pwd.encode(Encoding.default_internal)}\\#{one}".gsub(/\//, "\\")
				result = KWord.fetch_keys(file_path, @detect)
				result['file_name'] = one
				@all << result
			end
		}
		# pp @all
	end

	def out file
		puts "*** Output ***"
		# File.open(file, 'w'){|fout|
		# 	if @all.size>0
		# 		fout.puts @all[0].map{|k, v| k}.join(',') 
		# 		@all.each{|one|
		# 			fout.puts one.values.join(',')
		# 		}
		# 	end
		# }
		# rescue Exception => detail
		# 	puts "先关闭#{file}'!"
		# 	puts detail

		excel = CExcel2.new
		excel.open_rw('config.xlsx', file){|wb|
			sht = wb.worksheets(1)
			cells = @all.map{|one|
				one.sort.map{|k, v|
					v
				}
			}
			excel.write_area sht, 'A2', cells
		}

	end
end

def main
	detect = YAML::load_file('detect.yml')
	sum = SummaryParagraph.new(detect)
	sum.out 'Total.xlsx'
end

# ARGV = ['temp.doc']

if ARGV.size==0 #读结果并写入CSV文件
	main
elsif ARGV.size==1	#从模板取detect
	# read template
	pp keys = KWord.detect_keys("#{Dir.pwd}\\#{ARGV[0]}".gsub(/\//, "\\"))
	# pp keys.to_yaml.encoding
	open("detect.yml","w:utf-8") do |f|
		f.write keys.to_yaml
		# YAML.dump(keys, f)
	end
	puts "Write 'detect.yml' OK!"
else
	puts "Usage：CntWord.exe [Temp.doc]!"
end














