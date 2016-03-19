require 'docx_report/block_value'

module DocxReport
  class Table
    include BlockValue
    attr_reader :name, :has_header, :records

    def initialize(name, has_header = false, collection = nil, fields = [],
                   tables = [])
      @name = name
      @has_header = has_header
      @value = collection
      @records = []
      @fields = fields
      @tables = tables
    end

    def new_record
      new_record = Record.new
      @records << new_record
      new_record
    end

    def add_field(name, value = nil, type = :text, &block)
      field = Field.new(name, value || block, type)
      raise 'duplicate field name' if @fields.any? do |f|
        f.name == field.name
      end
      @fields << field
    end

    def add_table(name, collection = nil, has_header = false, &block)
      raise 'duplicate table name' if @tables.any? { |t| t.name == name }
      table = Table.new name, has_header, collection || block
      @tables << table
      table
    end

    def load_table(item)
      table = Table.new(name, has_header, nil, @fields, @tables)
      table.load_records(load_value(item))
      table
    end

    def load_records(collection)
      collection.each do |item|
        record = new_record
        @tables.each { |table| record.tables << table.load_table(item) }
        @fields.each { |field| record.fields << field.load_field(item) }
      end
    end
  end
end
