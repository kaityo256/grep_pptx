require 'tmpdir'
require 'rexml/document'

if ARGV.size < 1
  puts "usage: ruby grep_pptx.rb keyword"
  exit
end
keyword = ARGV[0]

def mygrep(dir, keyword, inputfile)
  Dir.glob(dir+"/ppt/slides/*.xml") do |file|
    slidenum = 0
    if file=~/slide([0-9]+).xml/
      slidenum = $1.to_i
    end
    f = open(file)
    doc = REXML::Document.new(f.gets(nil))
    n =  REXML::XPath.first(doc.root,'/p:sld/p:cSld')
    text = REXML::XPath.match(n, './/text()').join
    if text.include?(keyword)
      puts "find \"#{keyword}\" in #{inputfile} at slide #{slidenum}"
    end
  end
end

def find_in_file(inputfile, keyword)
  Dir.mktmpdir(nil,'./') do |dir|
    `cd #{dir};unzip ../#{inputfile}`
    mygrep(dir, keyword, inputfile)
  end
end

Dir.glob './**/*.pptx' do |file|
  find_in_file(file, keyword)
end
