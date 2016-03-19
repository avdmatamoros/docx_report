module DocxReport
  module BlockValue
    def load_value(item)
      if @value.is_a? Proc
        @value.call(item)
      else
        item.is_a?(Hash) ? item[@value] : item.send(@value)
      end
    end
  end
end
