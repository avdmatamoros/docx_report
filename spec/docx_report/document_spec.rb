require 'spec_helper'

describe DocxReport::Document do
  subject { DocxReport::Document.new 'spec/files/template.docx' }

  it 'loads content xml files' do
    expect(subject.files.keys).to match_array ['word/document.xml',
                                               'word/header.xml',
                                               'word/footer.xml']
  end

  it 'saves new docx file' do
    temp = Tempfile.new 'output.docx'
    subject.save temp.path
    expect(File.exists? temp.path).to be true
    temp.close!
  end
end
