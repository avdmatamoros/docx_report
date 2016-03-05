require 'spec_helper'

describe DocxReport::Document do
  subject { DocxReport::Document.new 'spec/files/template.docx' }

  it 'loads content xml files' do
    expect(subject.files.keys).to match_array ['word/document.xml',
                                               'word/header.xml',
                                               'word/footer.xml']
  end

  it 'saves new docx file' do
    subject.save 'output.docx'
    expect(File.exists? 'output.docx').to be true
    File.delete 'output.docx'
  end
end
