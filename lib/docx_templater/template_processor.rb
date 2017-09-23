require 'nokogiri'

module DocxTemplater
  class TemplateProcessor
    attr_reader :data, :escape_html, :rels_data

    # data is expected to be a hash of symbols => string or arrays of hashes.
    def initialize(data, escape_html = true, rels_data = [])
      @data = data
      @escape_html = escape_html
      @rels_data = rels_data
    end

    def render(document)
      document.force_encoding(Encoding::UTF_8) if document.respond_to?(:force_encoding)
      render_data(document, data)
    end

    def render_rels(document_rels)
      document_rels.force_encoding(Encoding::UTF_8) if document_rels.respond_to?(:force_encoding)
      render_rels_data(document_rels, rels_data)
    end

    private

    def render_rels_data(document_rels, rels_data)
      return document_rels unless rels_data.any?
      xml = Nokogiri::XML(document_rels)
      rels = xml.xpath('//r:Relationships', 'r': 'http://schemas.openxmlformats.org/package/2006/relationships').first
      unless rels.nil?
        rels_data.each do |rel|
          node = Nokogiri::XML::Node.new 'Relationship', xml
          node['Id'] = rel[:id]
          node['Type'] = rel[:type]
          node['Target'] = rel[:target]
          rels << node
        end
      end
      xml.to_s
    end

    def render_data(document, data)
      data.each do |key, value|
        document = render_value(document, key, value)
      end
      document
    end

    def render_value(document, key, value)
      case value
        when DocxTemplater::Block
          document = enter_block(document, key, value)
        when Array
          document = enter_multiple_values(document, key, value)
          document.gsub!("#SUM:#{key.to_s.upcase}#", value.count.to_s)
        when TrueClass, FalseClass
          if value
            document.gsub!(/\#(END)?IF:#{key.to_s.upcase}\#/, '')
          else
            document.gsub!(/\#IF:#{key.to_s.upcase}\#.*\#ENDIF:#{key.to_s.upcase}\#/m, '')
          end
        else
          document.gsub!("$#{key.to_s.upcase}$", safe(value))
      end
      document
    end

    def safe(text)
      if escape_html
        text.to_s.gsub('&', '&amp;').gsub('>', '&gt;').gsub('<', '&lt;')
      else
        text.to_s
      end
    end

    def enter_block(document, key, value)
      left_anchor = "<!--BEGIN_BLOCK_#{key.to_s.upcase}-->"
      right_anchor = "<!--END_BLOCK_#{key.to_s.upcase}-->"
      block_start = document.index(left_anchor)
      block_end = document.rindex(right_anchor)
      if !block_start.nil? && !block_end.nil?
        block_start += left_anchor.length
        left_part = document[0..block_start - 1]
        right_part = document[block_end + right_anchor.length..-1]
        block = document[block_start..block_end - 1]

        document = left_part
        value.context.each do |context|
          document += render_data(block.dup, context)
        end
        document += right_part
      end
      document
    end

    def enter_multiple_values(document, key, value)
      DocxTemplater.log("enter_multiple_values for: #{key}")
      if document.start_with?('<?xml')
        document_wrapped = false
      else
        document_wrapped = true
        document_wrapper = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14 wp14"><w:body>'
        document_wrapper += document
        document_wrapper += '</w:body></w:document>'
        document = document_wrapper
      end
      # TODO: ideally we would not re-parse xml doc every time
      xml = Nokogiri::XML(document)

      begin_row = "#BEGIN_ROW:#{key.to_s.upcase}#"
      end_row = "#END_ROW:#{key.to_s.upcase}#"
      begin_row_template = xml.xpath("//w:tr[contains(., '#{begin_row}')]", xml.root.namespaces).first
      end_row_template = xml.xpath("//w:tr[contains(., '#{end_row}')]", xml.root.namespaces).first
      DocxTemplater.log("begin_row_template: #{begin_row_template}")
      DocxTemplater.log("end_row_template: #{end_row_template}")
      raise "unmatched template markers: #{begin_row} nil: #{begin_row_template.nil?}, #{end_row} nil: #{end_row_template.nil?}. This could be because word broke up tags with it's own xml entries. See README." unless begin_row_template && end_row_template

      row_templates = []
      row = begin_row_template.next_sibling
      while row != end_row_template
        row_templates.unshift(row)
        row = row.next_sibling
      end
      DocxTemplater.log("row_templates: (#{row_templates.count}) #{row_templates.map(&:to_s).inspect}")

      # for each data, reversed so they come out in the right order
      value.reverse_each do |each_data|
        DocxTemplater.log("each_data: #{each_data.inspect}")

        # dup so we have new nodes to append
        row_templates.map(&:dup).each do |new_row|
          DocxTemplater.log("   new_row: #{new_row}")
          innards = new_row.inner_html
          matches = innards.scan(/\$EACH:([^\$]+)\$/)
          unless matches.empty?
            DocxTemplater.log("   matches: #{matches.inspect}")
            matches.map(&:first).each do |each_key|
              DocxTemplater.log("      each_key: #{each_key}")
              innards.gsub!("$EACH:#{each_key}$", safe(each_data[each_key.downcase.to_sym]))
            end
          end
          each_data.each do |(key, value)|
            next unless value.is_a?(DocxTemplater::Block)
            innards = render_value(innards.dup, key, value)
          end
          # change all the internals of the new node, even if we did not template
          new_row.inner_html = innards
          # DocxTemplater::log("new_row new innards: #{new_row.inner_html}")

          begin_row_template.add_next_sibling(new_row)
        end
      end
      (row_templates + [begin_row_template, end_row_template]).each(&:unlink)
      return xml.xpath('//w:body', xml.root.namespaces).first.inner_html if document_wrapped
      xml.to_s
    end
  end
end
