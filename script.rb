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
  attr_accessor :config, :pixabay_key, :ws, :time_stamp, :lesson_number,
                :deck, :course_name, :lesson_name, :words, :vocabulary_sheet


  def initialize
    set_config_vars
    connect_to_vocabulary_workbook
    @lesson_number = 100
    @course_name = "Test Course"
    @lesson_name = "Test Lesson"
    set_output_directory
    download
    create_deck
    write_deck_title
    create_slides_for_words
    save_deck
  end

  def create_slides_for_words
    @words.each do |word_hash|
      puts "Creating slides for #{word_hash[:main]}"
      create_three_slides_for_word(word_hash)
    end
  end

  def download
    (7..8).each do |row|
      if ws[row, 8] == @lesson_number.to_s
        word = {
          main: ws[row,1],
          part_of_speech: ws[row, 3],
          ipa: ws[row, 4],
          sentence: ws[row, 5],
          definition: ws[row, 6]
        }

        puts "searching: " +  word[:main]
        download_images_for_word(word)
      else
        puts "Finished all words for #{@lesson_number}"
        break
      end
    end
  end

  private

  def download_images_for_word(word_hash)
    search_term = word_hash[:main].gsub(" ", "+")
    url = "https://pixabay.com/api/?key=#{pixabay_key}&q=#{search_term}&image_type=photo"

    begin
      response = HTTParty.get(url)
      # Download the images for the entry
      response["hits"][0, 10].each_with_index do |hit, index|
        image_url = hit["webformatURL"]
        puts "GETTING: #{image_url}"
        download_image = open(image_url)
        FileUtils.mkdir_p(@directory) unless File.directory?(@directory)
        created_image_file_path = "#{@directory}/#{search_term} - #{index + 1}#{File.extname(image_url)}"
        IO.copy_stream(download_image, created_image_file_path)
        word_hash[:image_path] = created_image_file_path if index == 0
        puts "Successfully added image!"
      end

      rescue Exception => error
        puts "Error with #{word_hash[:main]}..."
        puts error.message
    end
    @words << word_hash

  end

  def set_config_vars
    @config = YAML.load_file('config.yaml')
    @pixabay_key = config["PIXABAY_KEY"]
    @words = []
    @vocabulary_sheet= @config["GOOGLE_SHEET_VOCABULARY"]
    @quiz_sheet = @config["GOOGLE_SHEET_QUIZ"]
  end

  def connect_to_vocabulary_workbook
    OpenSSL::SSL.const_set(:VERIFY_PEER, OpenSSL::SSL::VERIFY_NONE)
    session = GoogleDrive::Session.from_config("config.json")
    @ws = session.spreadsheet_by_key(@vocabulary_sheet).worksheets[0]
  end

  def set_output_directory
    @time_stamp = DateTime.now.strftime("%s")
    @directory = "output/#{@lesson_number} - #{@time_stamp}"
  end

  def create_deck
    @deck = Powerpoint::Presentation.new
  end

  def write_deck_title
    title = @course_name
    subtitle = @lesson_name
    @deck.add_intro title, subtitle
  end

  def create_three_slides_for_word(word_hash)
    create_word_slide(word_hash)
    create_detailed_word_slide(word_hash)
    create_image_slide(word_hash)
  end

  def create_word_slide(word)
    title = word[:main]
    content = [""]
    @deck.add_textual_slide title
  end

  def create_detailed_word_slide(word)
    title = word[:main]
    content = [word[:part_of_speech], word[:ipa], word[:definition], word[:sentence]]
    @deck.add_textual_slide title, content
  end

  def create_image_slide(word)
    title = word[:main]
    image_path = word[:image_path]
    if word[:image_path].length > 5
      @deck.add_pictorial_slide title, image_path
    end
  end

  def save_deck
    @deck.save("#{@directory}/#{@lesson_name}.pptx")
    puts "Deck Saved"
  end

end

PowerPointMaker.new

