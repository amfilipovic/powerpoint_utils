# Load the System.Windows.Forms assembly.
Add-Type -AssemblyName System.Windows.Forms

# Set the path to the current folder.
$path = (Get-Location).Path

# Create an instance of PowerPoint.
$powerpoint_app = New-Object -ComObject PowerPoint.Application

# Get all presentations in the current folder.
$presentations = Get-ChildItem -Path $path -Filter *.pptx

# Set the thumbnail size to 50% of the screen resolution.
$screen_width = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width / 2
$screen_height = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height / 2
$thumbnail_size = [Math]::Round($screen_width), [Math]::Round($screen_height)

# Initialize counters for the number of thumbnails and presentations processed.
$thumbnail_count = 0
$presentation_count = 0

# Loop through each presentation and extract its slides as PNG images.
foreach ($presentation in $presentations) {
    # Open the presentation.
    $presentation = $powerpoint_app.Presentations.Open($presentation.FullName)
    # Loop through each slide in the presentation.
    foreach ($slide in $presentation.Slides) {
        # Generate a unique filename for the thumbnail by appending a suffix to the original filename.
        $thumbnail_filename = $presentation.FullName.Replace(".pptx", "") + "_thumb_" + ($slide.SlideIndex).ToString("000") + ".png"
        # Export the slide as a PNG image with the specified thumbnail size.
        $slide.Export($thumbnail_filename, "png", $thumbnail_size[0], $thumbnail_size[1])
        # Increment the thumbnail count.
        $thumbnail_count += 1
    }
    # Close the presentation.
    $presentation.Close()
    # Increment the presentation count.
    $presentation_count += 1
}

# Quit PowerPoint.
$powerpoint_app.Quit()

# Print the number of thumbnails and presentations processed.
"Generated $thumbnail_count thumbnail(s) from $presentation_count presentation(s)."