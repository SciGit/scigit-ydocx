#!/usr/bin/env ruby
# encoding: utf-8

require 'ydocx'
require 'ydocx/differ'

module YDocx
  class Command
    class << self
      @@help    = /^\-(h|\-help)$/u
      @@version = /^\-(v|\-version)$/u
      def error(message='')
        puts message
        puts "see `#{self.command} --help`"
        exit
      end
      def extname(action)
        action == :to_html ? '.html': '.xml'
      end
      def command
        File.basename $0
      end
      def help
        banner = <<-BANNER
Usage: #{self.command} file [options]
  -h, --help      Display this help message.
  -v, --version   Show version.
        BANNER
        puts banner
        exit
      end
      def help_diff
        banner = <<-BANNER
Usage: #{self.command} file1 file2 output_file [options]
  -h, --help      Display this help message.
  -v, --version   Show version.
        BANNER
        puts banner
        exit
      end
      def report(action, path)
        puts "#{self.command}: generated #{File.expand_path(path)}"
      end
      def run(action=:to_html)
        argv = ARGV.dup
        if argv.empty? or argv[0] =~ @@help
          self.help
        elsif argv[0] =~ @@version
          self.version
        else
          file = argv.shift
          path = File.expand_path(file)
          if !File.exist?(path)
            self.error "#{self.command}: cannot open #{file}: No such file"
          elsif !File.extname(path).match(/^\.docx$/)
            self.error "#{self.command}: cannot open #{file}: Not a docx file"
          else
            doc = YDocx::Document.open(path)
            doc.send(action, path)
            ext = self.extname(action)
            self.report action, doc.output_file(ext[1..-1])
          end
        end
      end
      def run_diff
        argv = ARGV.dup
        if argv.empty? or argv[0] =~ @@help
          self.help_diff
        elsif argv[0] =~ @@version
          self.version
        elsif argv.length != 3
          self.help_diff
        else
          files = []
          argv[0..1].each do |file|
            path = File.expand_path(file)
            if !File.exist?(path)
              self.error "#{self.command}: cannot open #{file}: No such file"
            elsif !File.extname(path).match(/^\.docx$/)
              self.error "#{self.command}: cannot open #{file}: Not a docx file"
            else
              files << path
            end
          end
          STDOUT.sync = true
          puts 'Parsing...'
          docs = files.map { |f| YDocx::Document.open(f) }
          f = File.new(argv[2], "w")
          #require 'ruby-prof'
          #RubyProf.start
          t = Time.now
          f.write YDocx::Differ.new.diff(*docs)
          printf "Diff time: %f\n", Time.now - t
          #result = RubyProf.stop
          #printer = RubyProf::GraphHtmlPrinter.new(result)
          #printer.print(STDOUT)
          f.close
        end
      end
      def version
        puts "#{self.command}: version #{VERSION}"
        exit
      end
    end
  end
end
