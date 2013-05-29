#!/usr/bin/env ruby
# encoding: utf-8

require 'ydocx'
require 'diff/lcs'

module YDocx
  class Differ
    def self.get_text(chunk)
      if chunk.empty?
        ''
      elsif chunk[0].is_a? Image
        [chunk[0].img_hash]
      elsif chunk[0].is_a? Cell
        [chunk[0].blocks.hash]
      else
        chunk.join('')
      end
    end
    def self.chunk_similarity(chunk1, chunk2)
      text1 = chunk1.map { |c| get_text(c) }
      text2 = chunk2.map { |c| get_text(c) }
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
    def self.block_similarity(block1, block2)
      if block1.class != block2.class
        return 0
      end
      chunk_similarity(block1.get_chunks, block2.get_chunks)      
    end
    def self.get_chunks(paragraphs)
      chunks = []
      paragraphs.each_with_index do |p, i|
        if i > 0
          chunks << [Run.new("\r", Style.new)]
        end
        chunks += p.get_chunks
      end
      chunks
    end
    def self.diff(doc1, doc2)
      blocks1 = doc1.contents.blocks
      blocks2 = doc2.contents.blocks
      # Do an n^2 LCS diff on the blocks.
      n = blocks1.length
      m = blocks2.length
      lcs = Array.new(n+1) { Array.new(m+1, 0) }
      action = Array.new(n+1) { Array.new(m+1, -1) }
      blocks1.reverse.each_with_index do |a, ii|
        blocks2.reverse.each_with_index do |b, jj|
          if n*m > 1000
            sim = (a == b ? 1 : 0)
          else
            sim = block_similarity(a, b)
          end
          i = n-1-ii
          j = m-1-jj
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
            if lblocks.empty? || rblocks.empty? ||
               chunk_similarity(get_chunks(lblocks), get_chunks(rblocks)) > 0.5
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
        if lblocks.empty? || rblocks.empty? ||
           chunk_similarity(get_chunks(lblocks), get_chunks(rblocks)) > 0.5
          diff_blocks << [lblocks.dup, rblocks.dup]
        else
          diff_blocks << [lblocks.dup, []]
          diff_blocks << [[], rblocks.dup]
        end
      end
      
      table = Table.new
      diff_blocks.each do |block|
        row = [Cell.new, Cell.new]
        if block[0].empty?
          row[1].class = 'add'
          row[1].blocks = block[1]
        elsif block[1].empty?
          row[0].class = 'delete'
          row[0].blocks = block[0]
        elsif block[0] != block[1]
          row[0].class = row[1].class = 'modify'
          chunks = [get_chunks(block[0]), get_chunks(block[1])]
          changed = [Array.new(chunks[0].length), Array.new(chunks[1].length)]
          Diff::LCS.diff(chunks[0], chunks[1]).each do |diff|
            count = diff.map { |c| c.action }.uniq.length
            diff.each do |change|
              changed[change.action == '-' ? 0 : 1][change.position] = count
            end
          end
          
          for i in 0..1
            p = Paragraph.new
            group = RunGroup.new
            prev_table = nil
            chunks[i].each_with_index do |chunk, j|
              if chunk[0].is_a?(Cell)
                if changed[i][j]
                  if changed[i][j] == 2
                    chunk[0].class = 'modify'
                  elsif i == 0
                    chunk[0].class = 'delete'
                  else
                    chunk[0].class = 'add'
                  end
                end
                if chunk[0].parent != prev_table
                  prev_table = chunk[0].parent
                  row[i].blocks << prev_table
                end
              else
                if changed[i][j]
                  if changed[i][j] == 2
                    group.class = 'modify'
                  elsif i == 0
                    group.class = 'delete'
                  else
                    group.class = 'add'
                  end
                  if chunk == [Run.new("\n")]
                    group.runs << Run.new("&crarr;\n")
                  elsif chunk == [Run.new("\r")] 
                    group.runs << Run.new("&para;")
                    row[i].blocks << p
                    p.runs << group
                    group = RunGroup.new
                    p = Paragraph.new
                  else
                    group.runs += chunk
                  end
                else
                  p.runs << group unless group.runs.empty?
                  group = RunGroup.new
                  if chunk == [Run.new("\r")]
                    row[i].blocks << p
                    p = Paragraph.new
                  end
                  p.runs += chunk
                end
              end
            end
            p.runs << group unless group.runs.empty?
            row[i].blocks << p
          end
        else
          row[0].blocks = block[0]
          row[1].blocks = block[1]
        end
        table.cells << row
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