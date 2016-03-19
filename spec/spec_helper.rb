require 'codeclimate-test-reporter'
CodeClimate::TestReporter.start

$LOAD_PATH << File.join(File.dirname(__FILE__), '..', 'lib')
$LOAD_PATH << File.join(File.dirname(__FILE__))
require 'rspec'
require 'docx_report'

RSpec.configure do |config|
  config.color = true
end
