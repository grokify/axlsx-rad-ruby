require 'test/unit'
require 'axlsx_rad'

class AxlsxRadTest < Test::Unit::TestCase
  def testSetup

    config    = AxlsxRad::Config.new
    workbook  = AxlsxRad::Workbook.new
    worksheet = workbook.addWorksheet("My Sheet",config)

    assert_equal 'AxlsxRad::Config',    config.class.name
    assert_equal 'AxlsxRad::Workbook',  workbook.class.name
    assert_equal 'AxlsxRad::Worksheet', worksheet.class.name

  end
end