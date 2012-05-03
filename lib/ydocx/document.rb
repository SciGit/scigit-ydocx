#!/usr/bin/env ruby
# encoding: utf-8

require 'pathname'
require 'zip/zip'
require 'RMagick'
require 'ydocx/parser'
require 'ydocx/builder'

module YDocx
  class Document
    attr_reader :contents, :indecies, :pictures
    def self.open(file)
      self.new(file)
    end
    def initialize(file)
      @contents = nil
      @indecies = nil
      @pictures = []
      @path = nil
      @files = nil
      @zip = nil
      read(file)
    end
    def to_html(file='', options={})
      html = ''
      @files = @path.dirname.join(@path.basename('.docx').to_s + '_files')
      Builder.new(@contents) do |builder|
        builder.title = @path.basename
        builder.files = @files
        builder.style = options[:style] if options.has_key?(:style)
        if @indecies
          builder.indecies = @indecies
        end
        html = builder.build_html
      end
      unless file.empty?
        create_files if has_picture?
        html_file = @path.sub_ext('.html')
        File.open(html_file, 'w:utf-8') do |f|
          f.puts html
        end
      else
        html
      end
    end
    def to_xml(file='', options={})
      xml = ''
      Builder.new(@contents) do |builder|
        xml = builder.build_xml
      end
      unless file.empty?
        xml_file = @path.sub_ext('.xml')
        File.open(xml_file, 'w:utf-8') do |f|
          f.puts xml
        end
      else
        xml
      end
    end
    private
    def has_picture?
      !@pictures.empty?
    end
    def create_files
      FileUtils.mkdir @files unless @files.exist?
      @zip = Zip::ZipFile.open(@path.realpath)
      @pictures.each do |pic|
        origin_path = Pathname.new pic[:origin] # media/filename.ext
        source_path = Pathname.new pic[:source] # id/filename.ext
        dir = @files.join source_path.dirname
        FileUtils.mkdir dir unless dir.exist?
        binary = @zip.find_entry("word/#{origin_path}").get_input_stream
        if source_path.extname != origin_path.extname # convert
          image = Magick::Image.from_blob(binary.read).first
          image.format = source_path.extname[1..-1].upcase
          @files.join(source_path).open('w') do |f|
            f.puts image.to_blob
          end
        else
          @files.join(source_path).open('w') do |f|
            f.puts binary.read
          end
        end
      end
      @zip.close
    end
    def read(file)
      @path = Pathname.new file
      @zip = Zip::ZipFile.open(@path.realpath)
      doc = @zip.find_entry('word/document.xml').get_input_stream
      ref = @zip.find_entry('word/_rels/document.xml.rels').get_input_stream
      Parser.new(doc, ref) do |parser|
        @contents = parser.parse
        @indecies = parser.indecies
        @pictures = parser.pictures
      end
      @zip.close
    end
  end
end
