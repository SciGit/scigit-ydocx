PKG_FILES = FileList[
  '[a-zA-Z]*',
  'lib/**/*',
]

begin
  require 'jeweler'
  Jeweler::Tasks.new do |s|
    s.name = "ydocx"
    s.version = "1.2.3"
    s.author = "Yasuhiro Asaka, Zeno R.R. Davatz"
    s.email = "yasaka@ywesee.com, zdavatz@ywesee.com"
    s.homepage = "http://www.github.com/zdavatz/ydocx"
    s.platform = Gem::Platform::RUBY
    s.summary = "Convert docx files to html"
    s.files = PKG_FILES.to_a
    s.has_rdoc = false
    s.extra_rdoc_files = ["README.txt"]
  end
rescue LoadError
  puts "Jeweler not available. Install it with: sudo gem install jeweler"
end

Jeweler::GemcutterTasks.new
