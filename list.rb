# coding: UTF-8
Encoding.default_internal = "utf-8"
# Encoding.default_external = "gbk"

require 'yaml'
require 'fileutils'
require 'pp'

# Dir["*/*"].each{|one|
# 	puts one.encode('utf-8')
# }

all = []
Dir.glob("files/*"){ |file| 
 	if /^files\/(.*)/=~file
 		all<<$1
 	end
 	# all<<file.encode('gbk')
}

all1 = all.sort{|a, b| a.encode('gbk')<=>b.encode('gbk')}.map{|one|
	one
}

all2 = []
all1.each_with_index{|one, i|
	all2 << [one, "#{i+1}_#{one}", i+1]
}

all2.each{|one|
	puts one.join("\t")
	FileUtils.mv("files/#{one[0]}", "files/#{one[1]}")
}

