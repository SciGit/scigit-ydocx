#!/usr/bin/env ruby
# encoding: utf-8

require 'ydocx'

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
      def report(action, path)
        puts "#{self.command}: generated #{File.expand_path(path)}"
        exit
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
      def version
        puts "#{self.command}: version #{VERSION}"
        exit
      end
    end
  end
end
