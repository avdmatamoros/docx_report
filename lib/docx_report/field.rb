module DocxReport
  class Field
    attr_reader :name, :value, :type

    def initialize(name, value = nil, type = :text, &block)
      @name = "@#{name}@"
      @type = type
      set_value(value || block)
    end

    def set_value(value = nil, &block)
      @value = value || block
    end

    def load_value(item)
      Field.new(name[1..-2], if @value.is_a? Proc
                               @value.call(item)
                             else
                               item.is_a?(Hash) ? item[@value] : item.send(@value)
                             end, type)
    end
  end
end
