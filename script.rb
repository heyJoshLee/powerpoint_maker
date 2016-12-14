require "HTTParty"
require "pry"
require "google_drive"
require "yaml"
# require 'active_support/core_ext/hash/conversions'
require 'powerpoint'
require 'open-uri'
require "date"
require 'fileutils'

class PowerPointMaker
  attr_accessor :config, :pixabay_key, :ws, :time_stamp, :lesson

  def initialize
    set_config_vars
    connect_to_workbook
    @lesson = 100
    set_output_directory
    start
  end

  def start
    (4..4).each do |row|
      if ws[row, 8] == @lesson.to_s
        word = ws[row, 1]
        puts "searching " + word
        download_images_for_word(word)
      else
        puts "Finished all words for #{@lesson}"
        break
      end
    end
  end

  private

  def download_images_for_word(raw_word)
    search_term = raw_word.gsub(" ", "+")
    url = "https://pixabay.com/api/?key=#{pixabay_key}&q=#{search_term}&image_type=photo"

    begin
      response = HTTParty.get(url)
      # Download the images for the entry
      response["hits"][0, 10].each_with_index do |hit, index|
        image_url = hit["webformatURL"]
        puts "GETTING: #{image_url}"
        download_image = open(image_url)
        FileUtils.mkdir_p(@directory) unless File.directory?(@directory)
        IO.copy_stream(download_image, "#{@directory}/#{search_term} - #{index + 1}#{File.extname(image_url)}")
        puts "Successfully added image!"
      end

      rescue Exception => error
        puts "Error with #{raw_word}..."
        puts e.message
    end
  end

  def set_config_vars
    @config = YAML.load_file('config.yaml')
    @pixabay_key = config["PIXABAY_KEY"]
  end

  def connect_to_workbook
    OpenSSL::SSL.const_set(:VERIFY_PEER, OpenSSL::SSL::VERIFY_NONE)
    session = GoogleDrive::Session.from_config("config.json")
    @ws = session.spreadsheet_by_key(@config["GOOGLE_SHEET"]).worksheets[0]
  end

  def set_output_directory
    @time_stamp = DateTime.now.strftime("%s")
    @directory = "output/#{@lesson} - #{@time_stamp}"
  end
end

PowerPointMaker.new
puts "End"

# @deck = Powerpoint::Presentation.new

# # Creating an introduction slide:
# title = 'Bicycle Of the Mind'
# subtitle = 'created by Steve Jobs'
# @deck.add_intro title, subtitle

# # Creating a text-only slide:
# # Title must be a string.
# # Content must be an array of strings that will be displayed as bullet items.
# title = 'Why Mac?'
# content = ['Its cool!', 'Its light.']
# @deck.add_textual_slide title, content

# # Creating an image Slide:
# # It will contain a title as string.
# # and an embeded image
# # title = 'Everyone loves Macs:'
# # image_path = 'samples/images/sample_gif.gif'
# # @deck.add_pictorial_slide title, image_path

# # Specifying coordinates and image size for an embeded image.
# # x and y values define the position of the image on the slide.
# # cx and cy define the width and height of the image.
# # x, y, cx, cy are in points. Each pixel is 12700 points.
# # coordinates parameter is optional.
# # coords = {x: 124200, y: 3356451, cx: 2895600, cy: 1013460}
# # @deck.add_pictorial_slide title, image_path, coords

# # Saving the pptx file to the current directory.
# @deck.save('test.pptx')