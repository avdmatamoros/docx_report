module DocxReport
  class Table
    attr_accessor :name, :has_header, :records

    def initialize(name, has_header = false)
      @name = name
      @has_header = has_header
      @records = []
      @fields = []
    end

    def new_record
      new_record = Record.new
      records << new_record
      new_record
    end

    def add_field(name, attribute = nil, &block)
      @fields << { name: name, attribute: attribute, block: block }
    end

    def load_records(collection)
      collection.each do |item|
        record = new_record
        @fields.each do |field|
          if field[:block].nil?
            record.add_field field[:name], attribute_value(item, field)
          else
            record.add_field field[:name], field[:block].call(item)
          end
        end
      end
    end

    private

    def attribute_value(item, field)
      attribute_name = field[:attribute]
      item.is_a?(Hash) ? item[attribute_name] : item.send(attribute_name)
    end
  end
end
