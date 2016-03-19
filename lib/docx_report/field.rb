
require 'docx_report/block_value'

module DocxReport
  class Field
    include BlockValue
    attr_reader :name, :value, :type

    def initialize(name, value = nil, type = :text, &block)
      @name = "@#{name}@"
      @type = type
      set_value(value || block)
    end

    def set_value(value = nil, &block)
      @value = value || block
    end

    def load_field(item)
      Field.new(name[1..-2], load_value(item), type)
    end
  end
end
