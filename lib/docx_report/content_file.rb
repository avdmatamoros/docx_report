module DocxReport
  class ContentFile
    attr_reader :name, :xml, :rels_name, :rels_xml, :new_rels

    def initialize(name, zip)
      @name = name
      @xml = Nokogiri::XML(zip.read(name))
      @rels_name = "#{name.sub '/', '/_rels/'}.rels"
      @new_rels = false
      @rels_xml = Nokogiri::XML(if zip.entries.any? { |r| r.name == @rels_name }
                                  zip.read(@rels_name)
                                else
                                  new_rels_xml
                                end)
    end

    def new_uniqe_id
      (@rels_xml.xpath('//*[@Id]').map do |e|
        e[:Id][3..-1].to_i if e.name == 'Relationship'
      end.compact.max || 0) + 1
    end

    def new_rels_xml
      @new_rels = true
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships'\
      'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'\
      '</Relationships>'
    end

    def rels_has_items?
      @rels_xml.xpath('//*[@Id]').any?
    end
  end
end
