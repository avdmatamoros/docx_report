module DocxReport
  class Hyperlink
    attr_accessor :target, :file, :id

    def initialize(target, file, id = nil)
      @target = target
      @file = file
      if id.nil?
        generate_id
      else
        @id = id
        file.rels_xml.xpath("//*[@Id='#{id}']").first[:Target] = target
      end
    end

    private

    def generate_id
      @id = "rId#{file.new_uniqe_id}"
      file.rels_xml.children.first << hyperlink_rels
    end

    def hyperlink_rels
      Nokogiri::XML(
        format('<Relationship Id="%s" Type="http://schemas.openxmlformats'\
        '.org/officeDocument/2006/relationships/hyperlink" Target="%s" '\
        'TargetMode="External"/>', @id, @target)).children.first
    end
  end
end
