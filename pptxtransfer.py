import os
import sys
import pptx
import logging
import tempfile
from pptx.util import Inches
from gtts import gTTS
from moviepy.editor import ImageClip, concatenate_videoclips, AudioFileClip

# Configure logging for debugging and tracking script execution
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def print_instructions():
    # Display usage instructions for the script
    print("Instructions for using pptxtransfer.py:")
    print("- Ensure all dependencies (python-pptx, gTTS, moviepy) are installed.")
    print("- Run the script with python pptxtransfer.py (or python3 pptxtransfer.py). Follow the prompts to input the PowerPoint file path.")
    print("- The script will convert the PowerPoint into a video and save it to the specified output path.")

def check_dependencies():
    # Verify that all required dependencies are installed
    dependencies = {
        "pptx": "python-pptx",
        "gtts": "gTTS",
        "moviepy": "moviepy"
    }

    missing_dependencies = []

    for import_name, package_name in dependencies.items():
        try:
            __import__(import_name)
        except ImportError:
            missing_dependencies.append(package_name)
   
    # Inform the user of any missing dependencies and how to install them
    if missing_dependencies:
        print("The following dependencies are missing:")
        for dep in missing_dependencies:
            print(f"- {dep}")
        print("\nPlease install them using the following command:")
        print(f"pip install {' '.join(missing_dependencies)}")
        return False
    return True

def validate_file_path(file_path, expected_extension, check_exists=True):
    # Validate the existence and extension of a given file path
    if check_exists and not os.path.exists(file_path):
        raise FileNotFoundError(f"The file {file_path} does not exist.")
    if not file_path.endswith(expected_extension):
        raise ValueError(f"The file {file_path} does not have the expected {expected_extension} extension.")
    return True


def export_slide_as_image(slide, presentation):
    # Export a single PowerPoint slide as an image file
     with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as image_file:
        img = slide.get_image()  # Extract slide image
        img.save(image_file.name)  # Save image to temporary file
        return image_file.name

        # Save the slide image
        img.save(image_file.name)

        return image_file.name

def check_for_speaker_notes(presentation):
    # Check if any slide in the presentation has speaker notes
    for slide in presentation.slides:
        if slide.notes_slide and slide.notes_slide.notes_text_frame.text:
            return True
    return False

def get_user_input_for_skipping_slides():
    # Interact with the user to determine how to handle slides without speaker notes
    user_input = input("No slides with speaker notes found. Do you want to display all slides for a specified duration? (yes/no): ").strip().lower()
    if user_input == "yes":
        while True:
            try:
                slide_duration = int(input("Enter the number of seconds to display each slide: "))
                break  # Exit loop if input is a valid integer
            except ValueError:
                print("Please enter a valid integer for the number of seconds.")
        return False, slide_duration
    return True, 0

def pptx_to_video(pptx_path, output_path):
    # Main function to convert PowerPoint presentation to a video file
    logging.info("Starting the conversion process.")
    # Validate input and output file paths
    try:
        validate_file_path(pptx_path, '.pptx')
        validate_file_path(output_path, '.mp4', check_exists=False)
    except (FileNotFoundError, ValueError) as e:
        logging.error(f"File path validation error: {e}")
        sys.exit(1)

    # Exit if dependencies are missing
    if not check_dependencies():
        logging.error("Missing dependencies.")
        sys.exit(1)
    presentation = pptx.Presentation(pptx_path)   # Load the PowerPoint presentation
    has_speaker_notes = check_for_speaker_notes(presentation)     # Check for speaker notes
    skip_slides, slide_duration = (False, 5)  # Default values
    if not has_speaker_notes:
        skip_slides, slide_duration = get_user_input_for_skipping_slides()
    
    temp_audio_files, temp_image_files = []     # Lists to store temporary file paths

    # Process each slide in the presentation
    for i, slide in enumerate(presentation.slides):
        try:
            # Convert speaker notes to audio and export slides as images
            notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ''
            if notes:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.mp3') as audio_file:
                    gTTS(notes).save(audio_file.name)  # Convert notes to speech
                    temp_audio_files.append(audio_file.name)
            image_file_name = export_slide_as_image(slide, presentation)
            temp_image_files.append(image_file_name)
        except Exception as e:
            logging.error(f"Error processing slide {i}: {e}")


    # Create video clips from slides and corresponding audio
    logging.info("Creating video clips from slides.")
    video_clips = []
       for i, image_filename in enumerate(temp_image_files):
        try:
            img_clip = ImageClip(image_filename)  # Create an image clip from the slide image
            # If there's corresponding audio, set the clip duration to the audio length
            if i < len(temp_audio_files):
                audio_clip = AudioFileClip(temp_audio_files[i])
                img_clip = img_clip.set_duration(audio_clip.duration).set_audio(audio_clip)
            else:  # For slides without audio, use a default duration
                img_clip = img_clip.set_duration(slide_duration)
            video_clips.append(img_clip)  # Add the clip to the list of video clips
        except Exception as e:
            logging.error(f"Error creating video clip for slide {i}: {e}")

    # Concatenate all video clips into a single video file
    try:
        final_clip = concatenate_videoclips(video_clips)
        final_clip.write_videofile(output_path, fps=24)  # Save the final video
        logging.info(f"Output video saved to {output_path}")
    except Exception as e:
        logging.error(f"Error concatenating video clips: {e}")

    # Clean up temporary files
    for filename in temp_audio_files + temp_image_files:
        if os.path.exists(filename):
            os.remove(filename)  # Delete the temporary file

if __name__ == "__main__":
    pptx_to_video('example.pptx', 'output_video.mp4')  # Call main function with example file paths
