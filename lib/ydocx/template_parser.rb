#!/usr/bin/env ruby
# encoding: utf-8

require 'nokogiri'

module YDocx
  class TemplateParser
    OPTION_DEFAULTS = {
      :placeholder => 'None or N/A'
    }

    def initialize(doc, fields, options)
      @label_nodes = {}
      @label_root = {}
      @node_label = {}
      @doc = doc
      @fields = {}
      @options = options.merge(OPTION_DEFAULTS)
      fields.each do |field|
        if field[:type].is_a? Array
          store = field[:id][-1] == 's' ? (@fields[field[:id]] = {}) : @fields
          field[:type].each do |subfield|
            store[subfield[:id]] = subfield
          end
        else
          @fields[field[:id]] = field
        end
      end
    end

    def replace(data)
      doc = Nokogiri::XML.parse(@doc)
      merge_runs(doc)
      group_values(doc.at_xpath('//w:document//w:body'), data)
      replace_runs(doc.at_xpath('//w:document//w:body'), data)
      return doc
    end

    def get_value(val, field)
      if val.nil?
        return val
      elsif field[:type] == 'radio'
        return field[:options][val]
      elsif field[:type] == 'currency'
        return '$' + val
      else
        return val
      end
    end

    def lookup(str, data, data_index = nil)
      parts = str.split('.')
      if parts.length == 1
        if @fields[parts[0]].nil?
          return nil
        else
          return get_value(data[parts[0]], @fields[parts[0]])
        end
      else
        # only handle 2 parts for now
        if @fields[parts[0]].nil?
          return nil
        elsif @fields[parts[0]][:type].nil? && !data_index.nil?
          return data[parts[0]] && data[parts[0]][data_index] &&
                 get_value(data[parts[0]][data_index][parts[1]], @fields[parts[0]][parts[1]])
        elsif @fields[parts[0]][:type] == 'checkbox'
          if @fields[parts[0]][:options][parts[1]].nil?
            return nil
          else
            return (data[parts[0]] && data[parts[0]][parts[1]] ? '&#9632; ' : '&#9633;')  + ' ' +
                @fields[parts[0]][:options][parts[1]]
          end
        else
          return nil
        end
      end
    end

    def stringify(obj)
      if obj.nil?
        return @options[:placeholder]
      elsif obj.is_a? Hash
        # checkbox
        true_keys = obj.select { |k,v| v }.keys
        return true_keys.join(', ')
      else
        return obj.to_s
      end
    end

    def merge_runs(node)
      node.xpath('.//w:p').each do |p|
        cur_child = 0
        while cur_child < p.children.length
          r = p.children[cur_child]
          text = r.at_xpath('w:t')
          if text.nil?
            cur_child += 1
            next
          end

          prop = r.at_xpath('w:rPr')
          prop = prop ? prop.to_s : ''
          while cur_child + 1 < p.children.length
            next_r = p.children[cur_child + 1]
            next_prop = next_r.at_xpath('w:rPr')
            next_prop = next_prop ? next_prop.to_s : ''
            if next_r.name.start_with?('bookmark') || next_r.name == 'proofErr'
              next_r.remove # These are inserted randomly, can't really tell what they do
            elsif next_prop == prop && (next_text = next_r.at_xpath('w:t'))
              text.content += next_text.content
              if next_text['xml:space']
                text['xml:space'] = next_text['xml:space']
              end
              next_r.remove
            else
              break
            end
          end

          cur_child += 1
        end
      end
    end

    def group_values(node, data)
      node.xpath('.//w:r').each do |run|
        run.children.each do |t|
          if t.name == 't'
            t.content.scan(/%([0-9a-zA-Z_\-\.\[\]]+)%/).each do |match|
              parts = match[0].split('.')
              if parts.length > 1
                if @fields[parts[0]] && !@fields[parts[0]][:id]
                  (@label_nodes[parts[0]] ||= []) << run
                end
              end
            end
          end
        end
      end

      @label_nodes.each do |label, nodes|
        # Find lowest common ancestor
        if nodes.length == 1
          # Default to finding the row if it's in a table
          # or the last paragraph if it isn't
          node = nodes[0]
          first_p = nil
          first_tr = nil
          while first_tr == nil && node.name != 'body'
            if node.name == 'tr'
              first_tr ||= node
            elsif node.name == 'p'
              first_p ||= node
            end
            node = node.parent
          end
          if root = first_tr || first_p
            @label_root[label] = root
            @node_label[root] = label
          end
        else
          node_paths = nodes.map do |node|
            path = []
            cur_node = node.parent
            while cur_node.name != 'body'
              path << cur_node
              cur_node = cur_node.parent
            end
            path.reverse
          end

          low = 0
          high = node_paths.map { |p| p.length }.max - 1
          while low <= high
            mid = (low + high + 1) / 2
            if node_paths.all? { |path| path[mid] == node_paths[0][mid] }
              low = mid
              break if low == high
            else
              high = mid - 1
            end
          end

          if low == high
            node = node_paths[0][low]
            if @node_label[node].nil?
              @label_root[label] = node
              @node_label[node] = label
            end
          end
        end
      end
    end

    def replace_runs(node, data, data_index = nil)
      # process if statements
      if node.name == 'tbl'
        if tr = node.at_xpath('w:tr')
          if tc = tr.at_xpath('w:tc')
            if t = tc.at_xpath('.//w:t')
              t.content.scan(/%if ([0-9a-zA-Z_\-\.\[\]]+)%/).each do |match|
                value = lookup(match[0], data, data_index)
                if value.nil? || value.empty? || value.downcase == 'no'
                  node.remove
                  return
                end
              end
              t.inner_html = t.content.gsub(/%if ([0-9a-zA-Z_\-\.\[\]]+)%/, '')
            end
          end
        end
      end

      if node.name == 'p'
        if t = node.at_xpath('.//w:t')
          t.content.scan(/%if ([0-9a-zA-Z_\-\.\[\]]+)%/).each do |match|
            value = lookup(match[0], data, data_index)
            if value.nil? || value.empty? || value.downcase == 'no'
              node.remove
              return
            end
          end
          t.inner_html = t.content.gsub(/%if ([0-9a-zA-Z_\-\.\[\]]+)%/, '')
        end

        cur_child = 0
        while cur_child < node.children.length
          r = node.children[cur_child]
          text = r.at_xpath('w:t')
          if text.nil?
            cur_child += 1
            next
          end

          last_index = 0
          processed = ''
          text.content.to_enum(:scan, /%([0-9a-zA-Z_\-\.\[\]]+)%/).each do
            match = Regexp.last_match
            index = match.pre_match.length
            if index > last_index
              processed += text.content[last_index..index-1]
            end
            processed += stringify(lookup(match[1], data, data_index))
            last_index = index + match[0].length
          end
          processed += text.content[last_index..-1]
          text.inner_html = processed
          cur_child += 1
        end
        return
      end

      node.children.each do |child|
        if label = @node_label[child]
          if data[label].nil? || data[label].length == 0
            child.remove
          else
            next_child = child
            for i in 1..data[label].length-1
              next_child = next_child.add_next_sibling(child.clone)
              replace_runs(next_child, data, i)
            end
            replace_runs(child, data, 0)
          end
        else
          replace_runs(child, data, data_index)
        end
      end
    end
  end
end
