#!/usr/bin/env ruby
# encoding: utf-8

require 'pathname'
require 'zip'
require 'ydocx/parser'
require 'ydocx/template_parser'
require 'ydocx/builder'

module YDocx
  class Document
    attr_reader :builder, :contents, :images, :parser, :path
    def self.open(file, image_url = '')
      self.new(file, image_url)
    end
    def self.extract_document(file)
      begin
        path = Pathname.new file
        zip = Zip::File.open(path.realpath)
      rescue
        return nil
      end

      doc = zip.find_entry('word/document.xml').get_input_stream
      xml = Nokogiri::XML.parse(doc)
      doc.close
      zip.close

      xml
    end
    def self.list_sections(file)
      begin
        path = Pathname.new file
        zip = Zip::File.open(path.realpath)
      rescue
        return []
      end

      sections = TemplateParser.new(zip).list_sections
      zip.close

      sections
    end
    def self.fill_template(file, data, fields, output_file, options = {})
      begin
        path = Pathname.new file
        zip = Zip::File.open(path.realpath)
      rescue
        return nil
      end

      new_files = TemplateParser.new(zip, fields, options).replace(data)

      # Replace this document with another one.
      if new_files[:file]
        zip.close
        opts = {}
        if new_files[:section]
          opts[:extract_section] = new_files[:section]
        end
        fill_template path.dirname.join(new_files[:file]), data, fields, output_file, opts
        return
      end

      buffer = Zip::OutputStream.write_buffer do |out|
        zip.entries.each do |e|
          out.put_next_entry(e.name)
          if new_files[e.name]
            out.write new_files[e.name]
            new_files.delete(e.name)
          else
            out.write e.get_input_stream.read
          end
        end

        new_files.each do |name, data|
          out.put_next_entry(name)
          out.write data
        end
      end

      File.open(output_file, "wb") { |f| f.write(buffer.string) }
      zip.close
    end
    def initialize(file, image_url = '')
      @parser = nil
      @builder = nil
      @contents = nil
      @images = []
      @path = Pathname.new('.')
      @files = Pathname.new(file).dirname
      @image_url = image_url
      @zip = nil
      init
      read(file)
    end
    def init
    end
    def output_directory
      @files
    end
    def output_file(ext)
      @path.sub_ext(".#{ext.to_s}")
    end
    def to_html(output=false)
      html = ''
      @files = Pathname.new(@image_url.empty? ? '.' : @image_url)
      @path = Pathname.new(@path.basename)
      files = output_directory
      html = Builder.build_page(@contents)
      if output
        create_files if has_image?
        html_file = output_file(:html)
        File.open(html_file, 'w:utf-8') do |f|
          f.puts html
        end
      end
      html
    end
    def create_files
      files_dir = output_directory
      mkdir Pathname.new(files_dir) unless files_dir.exist?
      @images.each do |image|
        origin_path = Pathname.new image[:origin] # media/filename.ext
        source_path = Pathname.new image[:source] # images/filename.ext
        image_dir = files_dir.join source_path.dirname
        FileUtils.mkdir image_dir unless image_dir.exist?
        organize_image(origin_path, source_path, image[:data])
      end
    end
   private
    def organize_image(origin_path, source_path, data)
      if source_path.extname != origin_path.extname # convert
        output_file = output_directory.join(source_path)
        output_file.open('wb') do |f|
          f.puts data
        end
        if defined? Magick::Image
          image = Magick::Image.read(output_file).first
          image.format = source_path.extname[1..-1].upcase
          output_directory.join(source_path).open('wb') do |f|
            f.puts image.to_blob
          end
        end
      else
        output_directory.join(source_path).open('wb') do |f|
          f.puts data
        end
      end
    end
    def has_image?
      !@images.empty?
    end
    def read(file)
      @path = Pathname.new file
      begin
        @zip = Zip::File.open(@path.realpath)
      rescue
        # assume blank document
        @contents = ParsedDocument.new
        return
      end
      doc = @zip.find_entry('word/document.xml').get_input_stream
      rel = @zip.find_entry('word/_rels/document.xml.rels').get_input_stream
      rel_xml = Nokogiri::XML.parse(rel)
      rel_files = []
      rel_xml.xpath('/').children.each do |relat|
        relat.children.each do |r|
          if file = @zip.find_entry('word/' + r['Target'])
            rel_files << {
              :id => r['Id'],
              :type => r['Type'],
              :target => r['Target'],
              :stream => file.get_input_stream
            }
          end
        end
      end
      rel = @zip.find_entry('word/_rels/document.xml.rels').get_input_stream
      @parser = Parser.new(doc, rel, rel_files, @image_url) do |parser|
        @contents = parser.parse
        @images = parser.images
      end
      @zip.close
    end
    def mkdir(path)
      return if path.exist?
      parent = path.parent
      mkdir(parent)
      FileUtils.mkdir(path)
    end
  end
end
