require 'nokogiri'

module DocxReport
  class Parser
    def initialize(document)
      @document = document
    end

    def replace_all_fields(fields)
      @document.files.values.each do |xml_element|
        replace_element_fields(fields, xml_element)
      end
    end

    def fill_all_tables(tables)
      @document.files.values.each do |xml_element|
        fill_element_tables(tables, xml_element)
      end
    end

    private

    def search_for_text(name, parent_element)
      parent_element.xpath(".//*[contains(text(), '#{name}')]")
    end

    def replace_element_fields(fields, parent_element)
      fields.each do |key, value|
        search_for_text(key, parent_element).map do |element|
          element.content = element.content.gsub key, value
        end
      end
    end

    def find_table(name, parent_element)
      parent_element.xpath(".//w:tbl[//w:tblCaption[@w:val='#{name}']][1]").first
    end

    def find_row(table_element, row_number)
      table_element.xpath(".//w:tr[#{row_number}]").first
    end

    def fill_element_tables(tables, parent_element)
      tables.each do |table|
        tbl = find_table table.name, parent_element
        if tbl
          row_number = table.has_header ? 2 : 1
          tbl_row = find_row tbl, row_number

          if tbl_row
            table.records.each do |record|
              new_row = tbl_row.dup
              tbl_row.add_previous_sibling new_row
              replace_element_fields record.fields, new_row
            end
            tbl_row.remove
          end
        end
      end
    end
  end
end
