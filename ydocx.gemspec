# Generated by jeweler
# DO NOT EDIT THIS FILE DIRECTLY
# Instead, edit Jeweler::Tasks in Rakefile, and run 'rake gemspec'
# -*- encoding: utf-8 -*-

Gem::Specification.new do |s|
  s.name = "ydocx"
  s.version = "1.2.3"

  s.required_rubygems_version = Gem::Requirement.new(">= 0") if s.respond_to? :required_rubygems_version=
  s.authors = ["Yasuhiro Asaka, Zeno R.R. Davatz"]
  s.date = "2013-06-10"
  s.email = "yasaka@ywesee.com, zdavatz@ywesee.com"
  s.executables = ["diffx", "docx2html", "docx2xml"]
  s.extra_rdoc_files = [
    "README.txt"
  ]
  s.files = [
    "Gemfile",
    "LICENSE",
    "README.txt",
    "Rakefile",
    "lib/ydocx.rb",
    "lib/ydocx/builder.rb",
    "lib/ydocx/command.rb",
    "lib/ydocx/differ.rb",
    "lib/ydocx/document.rb",
    "lib/ydocx/markup_method.rb",
    "lib/ydocx/parser.rb"
  ]
  s.homepage = "http://www.github.com/zdavatz/ydocx"
  s.require_paths = ["lib"]
  s.rubygems_version = "2.0.3"
  s.summary = "Convert docx files to html"

  if s.respond_to? :specification_version then
    s.specification_version = 4

    if Gem::Version.new(Gem::VERSION) >= Gem::Version.new('1.2.0') then
      s.add_runtime_dependency(%q<nokogiri>, [">= 0"])
      s.add_runtime_dependency(%q<rubyzip>, [">= 0"])
      s.add_runtime_dependency(%q<htmlentities>, [">= 0"])
      s.add_runtime_dependency(%q<roman-numerals>, [">= 0"])
      s.add_runtime_dependency(%q<rmagick>, [">= 0"])
      s.add_runtime_dependency(%q<ruby-prof>, [">= 0"])
      s.add_runtime_dependency(%q<diff-lcs>, [">= 0"])
    else
      s.add_dependency(%q<nokogiri>, [">= 0"])
      s.add_dependency(%q<rubyzip>, [">= 0"])
      s.add_dependency(%q<htmlentities>, [">= 0"])
      s.add_dependency(%q<roman-numerals>, [">= 0"])
      s.add_dependency(%q<rmagick>, [">= 0"])
      s.add_dependency(%q<ruby-prof>, [">= 0"])
      s.add_dependency(%q<diff-lcs>, [">= 0"])
    end
  else
    s.add_dependency(%q<nokogiri>, [">= 0"])
    s.add_dependency(%q<rubyzip>, [">= 0"])
    s.add_dependency(%q<htmlentities>, [">= 0"])
    s.add_dependency(%q<roman-numerals>, [">= 0"])
    s.add_dependency(%q<rmagick>, [">= 0"])
    s.add_dependency(%q<ruby-prof>, [">= 0"])
    s.add_dependency(%q<diff-lcs>, [">= 0"])
  end
end
