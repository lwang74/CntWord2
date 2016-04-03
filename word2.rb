# encode: UTF-8
Encoding.default_internal = "utf-8"
# Encoding.default_external = "gbk"

require 'pp'
require 'win32ole'

class Word2
	def initialize doc_file
		@word = WIN32OLE.new('Word.Application')
		@doc = @word.Documents.open(doc_file, 'ReadOnly' => true)
	end

	def paragraphs 
		para = []
		@doc.paragraphs.each{|one|
			para << one.range.text.strip
		}
		para.map{|one|
			one if /\S+/=~one
		}.compact
	end

	def find_key
		find_list = []
		paragraphs.map{|one|
			if /\{(.+)\[(.+)\]\}/ =~one
				key = $2
				re = $1
				re += '\s*(.+)\s*'
				find_list<<{key => Regexp.new(re.gsub(/\s+/, '\s*'))}
			end
		}

		# p find_list.size
		# find_list.each{|one|
		# 	one.each{|k, v|
		# 		puts k
		# 		pp v
		# 	}
		# }
		find_list
	end

	def find keys
		val = {}
		keys.each{|one|
			one.each{|k, v|
				val[k] = nil 
			}
		}

		paragraphs.each{|one|
			keys.each{|keys|
				keys.each{|k, v|
					if v =~ one
						val[k] = $1
					end
				}
			}
		}
		val
	end

	def get_content doc_file
		@doc = @word.Documents.open(doc_file, 'ReadOnly' => true)
		# show_ole @doc.paragraphs
		# show_ole_methods @doc.paragraphs, 'ole_func_methods'
		
		# p @doc.ole_put_methods
		# show_ole_methods @doc.paragraphs
		find_list = []
		@doc.paragraphs.each{|one|
			# show_ole_methods one.range, 'ole_get_methods'
			# puts one.range.text.encoding
			if /\{(.+)\[(.+)\]\}/ =~one.range.text
				key = $2
				re = $1
				find_list<<{key => Regexp.new(re.gsub(/\s+/, '\s*'))}
			end
		}

		# p find_list.size
		find_list.each{|one|
			one.each{|k, v|
				puts k
				pp v
			}
		}
		
		@doc.close
	end
	
	def close
		@word.quit
	end
	
	protected
	def trim val
		val.gsub(/[\s\r\a\?]/, '').gsub(/[,]/, '，')
	end

	def show_ole obj
		p obj.methods.map{|one|
			one if /ole.*/=~one
		}.compact
	end 
	def show_ole_methods obj, method
		# p obj.send(method).map{|one| one.to_s}.sort
		obj.send(method).map{|one| one.to_s}.sort.uniq.each{|x|
			puts x
		}
	end
end

def dump result
	result.each{|k, v|
		puts "#{k}: #{v}"
	}
end

class KWord
	def self.detect_keys temp_file
		doc = Word2.new temp_file
		keys = doc.find_key 
		doc.close
		keys
	end 	

	def self.fetch_keys doc_file, keys
		doc = Word2.new doc_file
		result = doc.find keys 
		doc.close
		result
	end 	
end

if __FILE__==$0
	# temp = 'E:\Lyx1\2016_16届河东区学术年会\CntWord3\temp.doc'
	# a = "D:\\Lyx\\论文统计\\曹悦芊在教学实践中探索信息技术与历史学科课程整合的策略登记表.doc"
	# b = "D:\\Lyx\\论文统计\\张丽萍英语教学与信息技术的运用登记表.docx"
	# c = "D:\\Lyx\\论文统计\\陈寅论文登记表.docx"

	temp = 'E:\Lyx1\2016_16届河东区学术年会\CntWord3\temp.doc'
	c = 'E:\Lyx1\2016_16届河东区学术年会\CntWord3\登记表\李玉琴让历史课堂活起来登记表1.doc'
	# doc = Word2.new a
	# keys = doc.find_key 
	# pp keys
	# doc.close
	
	# doc = Word2.new c
	# para = doc.find keys
	# pp para
	# doc.close

	keys = KWord.detect_keys(temp)
	# pp KWord.fetch_keys(a, keys)
	# pp KWord.fetch_keys(b, keys)
	pp KWord.fetch_keys(c, keys)
end


