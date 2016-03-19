require 'docx_report/data_item'

module DocxReport
  class Report
    include DataItem
    attr_reader :fields, :tables

    def initialize(template_path)
      @template_path = template_path
      @fields = []
      @tables = []
    end

    def add_table(name, collection = nil, has_header = false)
      raise 'duplicate table name' if @tables.any? { |t| t.name == name }
      table = Table.new name, has_header
      @tables << table
      yield table
      table.load_records collection if collection
    end

    def generate_docx(filename = nil, template_path = nil)
      document = Document.new template_path || @template_path
      apply_changes document
      if filename.nil?
        document.save_to_memory
      else
        document.save_to_file filename
      end
    end

    private

    def apply_changes(document)
      parser = Parser.new document
      parser.replace_all_fields @fields
      parser.fill_all_tables @tables
    end
  end
end
