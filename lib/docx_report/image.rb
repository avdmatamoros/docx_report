require 'mini_magick'

module DocxReport
  class Image
    attr_accessor :path, :id, :nodes, :files, :new_rels, :type

    def initialize(path, document)
      @path = path
      @files = {}
      @nodes = []
      @new_rels = []
      global_id(document)
    end

    def global_id(document)
      @id = document.image_ids.max + 1
      document.image_ids << @id
      document.images << self
    end

    def file_image_id(file)
      new_image_id(file) if files[file.name].nil?
      files[file.name]
    end

    def image_ref(id)
      rels = Nokogiri::XML(
        format('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'\
        'officeDocument/2006/relationships/image" Target="media/image%s.%s"/>',
               id, @id, '%s')).children.first
      @new_rels << rels
      rels
    end

    def save(output)
      img = MiniMagick::Image.open(@path)
      @type = img.type.downcase
      fix_rels
      output.put_next_entry "word/media/image#{@id}.#{@type}"
      img.write output
      set_dimentions img.width, img.height, img.resolution
    end

    private

    def fix_rels
      new_rels.each { |rels| format(rels, @id, @type) }
    end

    def new_image_id(file)
      files[file.name] = "rId#{file.new_uniqe_id}"
      file.rels_xml.children.first << image_ref(files[file.name])
    end

    def set_dimentions(width, height, resolution)
      @width = width.to_f / resolution[0] * EMUS_PER_INCH
      @height = height.to_f / vert_dpi(resolution) * EMUS_PER_INCH
      fit_in_page
      @nodes.each do |node|
        node.xpath('.//*[@cx and @cy]').each do |ext|
          ext[:cx] = @width.to_i
          ext[:cy] = @height.to_i
        end
      end
    end

    def vert_dpi(resolution)
      resolution[resolution.length == 2 ? 1 : 2]
    end

    def fit_in_page
      if @width > MAX_WIDTH_EMUS
        ratio = @height / @width
        @width = MAX_WIDTH_EMUS
        @height = MAX_WIDTH_EMUS * ratio
      end
      if @height > MAX_HEIGHT_EMUS
        ratio = @width / @height
        @height = MAX_HEIGHT_EMUS
        @width = MAX_HEIGHT_EMUS * ratio
      end
    end

    EMUS_PER_INCH = 914400
    MAX_WIDTH_EMUS = 5742360
    MAX_HEIGHT_EMUS = 8229600
  end
end
