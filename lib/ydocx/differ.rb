#!/usr/bin/env ruby
# encoding: utf-8

require 'ydocx'
require 'diff/lcs'

module YDocx
  class Differ
    def initialize
      @cur_change_id = 1
    end
    def get_text(chunk)
      if chunk.empty?
        ''
      elsif chunk[0].is_a? Image
        [chunk[0].img_hash]
      elsif chunk[0].is_a? Cell
        chunk[0].blocks.map { |b| get_text(b.get_chunks) }
      else
        chunk.join('')
      end
    end
    def text_similarity(text1, text2)
      # Find the LCS between the runs
      lcs = Diff::LCS.LCS(text1, text2)
      lcs_len = 0
      lcs.each do |run|
        lcs_len += run.length
      end
      tlen = 0
      (text1 + text2).each do |run|
        tlen += run.length
      end
      if tlen == 0
        return 1
      else
        return 2.0 * lcs_len / tlen
      end
    end
    # Return the % similarity between two blocks (paragraphs, tables)
    def block_similarity(block1, text1, block2, text2)
      if block1.class != block2.class
        return 0
      end
      text_similarity(text1, text2)
    end
    def get_detail_blocks(blocks1, blocks2)
      n = blocks1.length
      m = blocks2.length
      if n == 0 || m == 0
        return [[blocks1, blocks2]]
      end
      if n*m > 10000
        # assume it's all different QQ
        return [[blocks1, []], [[], blocks2]]
      end
      lcs = Array.new(n+1) { Array.new(m+1, 0) }
      action = Array.new(n+1) { Array.new(m+1, -1) }
      text1 = blocks1.map { |b| b.get_chunks.map(&method(:get_text)) }
      text2 = blocks2.map { |b| b.get_chunks.map(&method(:get_text)) }
      blocks1.reverse.each_with_index do |a, ii|
        blocks2.reverse.each_with_index do |b, jj|
          i = n-1-ii
          j = m-1-jj
          sim = block_similarity(a, text1[i], b, text2[j])
          # printf "%d %d = %d\n", i, j, sim
          lcs[i][j] = lcs[i+1][j]
          action[i][j] = 0
          if lcs[i][j+1] > lcs[i][j]
            lcs[i][j] = lcs[i][j+1]
            action[i][j] = 1
          end
          if sim > 0.5 && lcs[i+1][j+1] + sim > lcs[i][j]
            lcs[i][j] = lcs[i+1][j+1] + sim
            action[i][j] = 2
          end
        end
      end
      
      i = 0
      j = 0
      lblocks = []
      rblocks = []
      diff_blocks = []
      while i < n || j < m        
        if j == m || action[i][j] == 0
          lblocks << blocks1[i]
          i += 1
        elsif i == n || action[i][j] == 1
          rblocks << blocks2[j]
          j += 1
        else
          unless lblocks.empty? && rblocks.empty?
            if lblocks.empty? || rblocks.empty?
              diff_blocks << [lblocks.dup, rblocks.dup]
            else
              diff_blocks << [lblocks.dup, []]
              diff_blocks << [[], rblocks.dup]
            end
            lblocks = []
            rblocks = []
          end
          diff_blocks << [[blocks1[i]], [blocks2[j]]]
          i += 1
          j += 1
        end
      end
      unless lblocks.empty? && rblocks.empty?
        if lblocks.empty? || rblocks.empty?
          diff_blocks << [lblocks.dup, rblocks.dup]
        else
          diff_blocks << [lblocks.dup, []]
          diff_blocks << [[], rblocks.dup]
        end
      end
      diff_blocks
    end
    def get_diff_change_array(list1, list2)
      change_id = [Array.new(list1.length), Array.new(list2.length)]
      Diff::LCS.diff(list1, list2).each do |diff|
        count = diff.map { |c| c.action }.uniq.length
        if count == 1
          cid = -1
        else
          cid = @cur_change_id
          @cur_change_id += 1
        end
        diff.each do |change|
          change_id[change.action == '-' ? 0 : 1][change.position] = cid
        end
      end
      change_id
    end
    def get_paragraph_diff(p1, p2, result)
      chunks = [p1.get_chunks, p2.get_chunks]
      change_id = get_diff_change_array(chunks[0], chunks[1])
      
      for i in 0..1
        p = Paragraph.new
        p.align = [p1,p2][i].align
        group = RunGroup.new
        prev_table = nil
        chunks[i].each_with_index do |chunk, j|
          if change_id[i][j]
            if change_id[i][j] >= 1
              group.class = (sprintf 'modify modify%d', change_id[i][j])
            elsif i == 0
              group.class = 'delete'
            else
              group.class = 'add'
            end
            if chunk == [Run.new("\n")]
              group.runs << Run.new("&crarr;\n")
            else
              group.runs += chunk
            end
          else
            p.runs << group unless group.runs.empty?
            group = RunGroup.new
            p.runs += chunk
          end
        end
        p.runs << group unless group.runs.empty?
        result[i].blocks << p
      end
    end
    def get_table_diff(t1, t2, result)
      tables = [Table.new, Table.new]
      get_detail_blocks(t1.rows, t2.rows).each do |rblock|
        if rblock[0].empty?
          rblock[1].map { |r| r.cells.map { |c| c.class = 'add' } }
          tables[1].rows += rblock[1]
        elsif rblock[1].empty?
          rblock[0].map { |r| r.cells.map { |c| c.class = 'delete' } }
          tables[0].rows += rblock[0]
        else # should be 1 row in each
          change_id = get_diff_change_array(rblock[0].first.cells, rblock[1].first.cells)
          for i in 0..1
            rblock[i].first.cells.each_with_index do |cell, j|
              if change_id[i][j]
                if change_id[i][j] >= 1
                  cell.class = (sprintf 'modify modify%d', change_id[i][j])
                elsif i == 0
                  cell.class = 'delete'
                else
                  cell.class = 'add'
                end
              end
            end
          end
          tables[0].rows += rblock[0]
          tables[1].rows += rblock[1]
        end
      end
      result[0].blocks << tables[0]
      result[1].blocks << tables[1]
    end
    def diff(doc1, doc2)
      blocks1 = doc1.contents.blocks
      blocks2 = doc2.contents.blocks
      
      puts 'Extracting text...'
      text1 = blocks1.map { |b| b.get_chunks.map(&method(:get_text)) }
      text2 = blocks2.map { |b| b.get_chunks.map(&method(:get_text)) }
      
      puts 'Computing paragraph diffs...'
      lblocks = []
      rblocks = []
      diff_blocks = []
      Diff::LCS.sdiff(text1.map(&:hash), text2.map(&:hash)).each do |change|
        if change.action == '='
          diff_blocks << [lblocks.dup, rblocks.dup] unless lblocks.empty? && rblocks.empty?
          diff_blocks << [[blocks1[change.old_position]], [blocks2[change.new_position]]]
          lblocks = []
          rblocks = []
        else
          lblocks << blocks1[change.old_position] unless change.old_element.nil?
          rblocks << blocks2[change.new_position] unless change.new_element.nil?
        end
      end
      diff_blocks << [lblocks.dup, rblocks.dup] unless lblocks.empty? && rblocks.empty?
      
      puts 'Computing block diffs...'
      table = Table.new
      diff_blocks.each do |dblock|
        get_detail_blocks(dblock[0], dblock[1]).each do |block|
          row = Row.new
          row.cells = [Cell.new, Cell.new]
          if block[0].empty?
            row.cells[1].class = 'add'
            row.cells[1].blocks = block[1]
          elsif block[1].empty?
            row.cells[0].class = 'delete'
            row.cells[0].blocks = block[0]
          elsif block[0] != block[1] # should only be 1 block in each
            row.cells[0].class = row.cells[1].class = 'modify'
            if block[0].first.is_a? Paragraph
              get_paragraph_diff(block[0].first, block[1].first, row.cells)
            else
              get_table_diff(block[0].first, block[1].first, row.cells)
            end
          else
            row.cells[0].blocks = block[0]
            row.cells[1].blocks = block[1]
          end
          table.rows << row
        end
      end
      
      [doc1, doc2].each do |doc|
        if !doc.images.empty?
          doc.create_files
        end
      end
      
      html_doc = ParsedDocument.new
      html_doc.blocks << table
      builder = Builder.new(html_doc)
      builder.title = 'Diff Results'
      builder.style = true
      builder.build_html
    end
  end
end