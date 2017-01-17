require "HTTParty"
require "pry"
require "google_drive"
require "yaml"
require 'powerpoint'
require 'open-uri'
require "date"
require 'fileutils'

class PowerPointMaker
  @@lessons = []
  attr_accessor :config, :pixabay_key, :ws, :time_stamp, :lesson_number,
                :deck, :course_name, :lesson_name, :words, :vocabulary_sheet,
                :quiz_questions, :session, :quiz_lesson_number


  def self.prompt_user_for_inputs
    process_type = ""

    until ["s", "m"].include?(process_type)
      puts "Do you want to process a single presentation or multiple presentations?"
      puts "[s/m]"
      process_type = gets.chomp.downcase
    end

    if process_type == "s"
      self.create_single_presentation
    elsif process_type == "m"
      self.create_multiple_presentations
    end
  end

  def initialize(course_name, lesson_name, lesson_number, quiz_lesson_number)
    @course_name = course_name
    @lesson_name = lesson_name
    @lesson_number = lesson_number
    @quiz_lesson_number = quiz_lesson_number
    set_config_vars
    connect_to_vocabulary_workbook
    # connect_to_quiz_workbook
    # create_quiz_questions_hash
    set_output_directory
    download
    create_deck
    write_deck_title
    create_slides_for_words
    # create_slides_for_questions
    save_deck
  end

  def self.create_single_presentation
    confirm = ""
      while confirm != "y"

        puts "Creating single presentation"
        
        puts "What's the course name?"
        course_name = gets.chomp
        
        puts "What the lesson name?"
        lesson_name = gets.chomp

        puts "What's the lesson number on the spreadsheet?"
        lesson_number = gets.chomp
        
        puts "What's the quiz lesson number on the spreadsheet?"
        quiz_lesson_number = gets.chomp

        puts "Is this correct?"
        puts "----------------"

        puts <<-CONFIRM_INPUTS
            Course Name: #{course_name}
            Lesson Name: #{lesson_name}
            Lesson Number: #{lesson_number}
            Quiz Lesson Number: #{quiz_lesson_number}
        CONFIRM_INPUTS

        puts "[y/n]"

        confirm = gets.chomp.downcase
      end

      PowerPointMaker.new(course_name, lesson_name, lesson_number, quiz_lesson_number)
  end

  def self.create_multiple_presentations
    puts "Creating multiple presentations."

    puts "Put the course name to be used for all presentations."
    course_name_for_all_lessons = gets.chomp.strip

    counter = 1
    
    auto_generate = ""
    puts "Auto Generate Lesson info? Only do this if you know what you're doing. [y/n]"
    auto_generate = gets.chomp.downcase.strip

    puts "What's the lesson type: Ex: Lesson, Episode, Chapter ('Lesson' is the default)"
    lesson_type = gets.chomp
    lesson_type = "Lesson" if lesson_type == ""


    # Mass create ppt by using a range, or comma separated list of numbers
    if auto_generate == "y"

      puts "Enter the lesson numbers as integers separated by commas. ex 1,2,5,9 OR as a range [1-20]"
      auto_generate_input = gets.chomp.strip

      # parse range or comma separated lessons
      if auto_generate_input.include?("-")
        min_and_max = auto_generate_input.split("-").map! { |m| m.to_i }
        lessons_numbers_to_generate_automatically = (min_and_max[0]..min_and_max[1]).to_a
      else
        lessons_numbers_to_generate_automatically = auto_generate_input.split(",")
      end

      # Create lesson params and put into @@lessons
      lessons_numbers_to_generate_automatically.each do |lesson_number|
        @@lessons << [course_name_for_all_lessons, "#{lesson_type} #{lesson_number}", (lesson_number.to_i * 1).to_s, lesson_number.to_s ]
      end
    else
      # Continue to prompt use for Lesson info
      add_another_lesson = ""
      until add_another_lesson == "n"
        add_lesson_to_lessons_array(counter, course_name_for_all_lessons)
        puts "Add another lesson? [y/n]"
        add_another_lesson = gets.chomp.downcase
        counter += 1
      end
    end


    # Generate all lessons
    puts "Start generating presentations."
    @@lessons.each.each_with_index do |l, index|
      puts "Creating lesson: #{l[1]}"
      PowerPointMaker.new(l[0], l[1], l[2], l[3])
      puts "Finished creating presentation for lesson."
      puts "#{@@lessons.count - (index + 1)} lessons left"
    end
  end

  def self.add_lesson_to_lessons_array(counter, course_name)
    # Allow user to compose an array that contains info for all lessons to  be processed
    confirm_inputs = ""
    until confirm_inputs == "y"

      puts "Presentation Number #{counter}:"
      puts "Enter the following. Make sure to separate with commas."
      puts "lesson_name, lesson_number, quiz_lesson_number"
      presentation_inputs = gets.chomp.split(",")
      presentation_inputs.map! { |a| a.strip }
      presentation_inputs.unshift(course_name)

      puts "Is this correct?"
      puts "----------------"

      puts <<-CONFIRM_INPUTS
          Course Name: #{presentation_inputs[0]}
          Lesson Name: #{presentation_inputs[1]}
          Lesson Number: #{presentation_inputs[2]}
          Quiz Lesson Number: #{presentation_inputs[3]}
      CONFIRM_INPUTS
      puts "[y/n]"

      confirm_inputs = gets.chomp.downcase
      @@lessons << presentation_inputs if confirm_inputs == "y"
    end
  end

  def create_slides_for_words
    @words.each do |word_hash|
      puts "Creating slides for #{word_hash[:main]}"
      create_three_slides_for_word(word_hash)
    end
  end

  def download
    started = false
    (2..@ws.num_rows).each do |row|
      unless started
        started = ws[row, 8] == @lesson_number
        next unless started
      end
      if ws[row, 8] == @lesson_number
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
    OpenSSL::SSL.const_set(:VERIFY_PEER, OpenSSL::SSL::VERIFY_NONE)
    @session = GoogleDrive::Session.from_config("config.json")
  end

  def connect_to_vocabulary_workbook
    @ws = @session.spreadsheet_by_key(@vocabulary_sheet).worksheets[0]
  end

  def connect_to_quiz_workbook
    begin
      @ws_2 = @session.spreadsheet_by_key(@quiz_sheet).worksheets[0]
    rescue Exception => error
      @ws_2 = false
    end
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
    if !word[:image_path].nil? && word[:image_path].length > 5
      @deck.add_pictorial_slide title, image_path
    end
  end

  def save_deck
    @deck.save("#{@directory}/#{@lesson_name}.pptx")
    puts "Deck Saved"
  end

  def create_quiz_questions_hash
    @quiz_questions = []
    question = {}
    started = false
    if @ws_2
      (2..@ws_2.num_rows).each_with_index do |row, index|
         unless started
              started = @ws_2[row, 5] == @quiz_lesson_number
            next unless started
          end
        
        if @ws_2[row,1] == "question"

         
          @quiz_questions << question unless question.empty?


          break if @ws_2[row, 5] != @quiz_lesson_number && @ws_2[row, 5] != ""

          question = {
            body: @ws_2[row, 3],
            choices: []
          }
        else
          choice = {
            body: @ws_2[row, 3],
            correct: (@ws_2[row, 4] == "TRUE" ? true : false)
          }
          question[:choices] << choice
        end
      end
    end
  end

  def create_slides_for_questions
    # if there are no quiz questions, don't create slides for the quizzes
    # create a question and answer slide for each question
    unless @quiz_questions.empty?
      @quiz_questions.each do |question|
        begin
          title = question[:body]
          answer_texts = question[:choices].map { |c| c[:body] }
          correct_answer = question[:choices].select { |c| c[:correct] == true}
          correct_answer_text = correct_answer[0][:body]
          content = answer_texts
          @deck.add_textual_slide title, content
          @deck.add_textual_slide title, [correct_answer_text]

          rescue Exception => error
                  puts "Error with #{correct_answer[0]}..."
                  puts error.message

        end
      end
    end
  end

end

PowerPointMaker.prompt_user_for_inputs