require 'spec_helper'

describe DocxReport::Document do
  subject { DocxReport::Document.new 'spec/files/template.docx' }

  it 'loads content xml files' do
    expect(subject.files.map(&:name)).to match_array ['word/document.xml',
                                                      'word/header1.xml',
                                                      'word/footer1.xml']
  end

  it 'saves new docx file' do
    temp = Tempfile.new 'output.docx'
    subject.save_to_file temp.path
    expect(File.exist?(temp.path)).to be true
    temp.close!
  end

  it 'saves new docx to memory' do
    expect(subject.save_to_memory).not_to be nil
  end
end
