require 'spec_helper'

describe DocxReport::Record do
  subject { DocxReport::Record.new }

  it 'adds text fields' do
    subject.add_field 'name', 'Ahmed Abudaqqa'
    expect(subject.fields).to eq('{@name}' => 'Ahmed Abudaqqa')
  end
end
