from PIL import Image, ImageFont, ImageDraw
import pandas as pd
import os
# Set the fonts and font sizes
name_font = ImageFont.truetype("AdventPro-SemiBold.ttf", 32)
content_font = ImageFont.truetype("AdventPro-SemiBold.ttf", 32)

# Read the data from the Excel file
data = pd.read_excel("reciept.xlsx")

# Loop through each row of data
for index, row in data.iterrows():

    # Load the certificate template image
    img = Image.open("image.jpg")

    # Create a ImageDraw object
    draw = ImageDraw.Draw(img)

    # Set the content to be printed on the certificate
    name = row['name']
    price = row['price']
    count = row['count']
    date_of_feeding = row['date_of_feeding']
    date_of_donation = row['date_of_donation']
    reciept_id = row['reciept_id']

    content = f"We sincerely like to thank you for your generous gift to our organization \n" \
              f"of Rs {price} We are thrilled to have your support.\n\n" \
              f"Through your contribution, we will feed {count} needy people through \n" \
              f"gifting meal boxes on {date_of_feeding} You have helped achieve our goal \n"\
              f"towards zero hunger.\n"\
              f"You truly made a difference today."

    # Set the starting position to print the name
    x = 485
    y = 320

    # Print the name with a different color
    draw.text((x, y), f"Dear ", font=content_font, fill=(0, 0, 0))
    x += content_font.getbbox("Dear ")[2]
    draw.text((x, y), name, font=name_font, fill=(255, 0, 0))
    x += name_font.getbbox(name)[2]

    # Set the starting position to print the content
    x = 485
    y = 400

    draw.text((175, 250), "DATE: ", font=content_font, fill=(255, 255, 255))
    draw.text((175 + content_font.getsize("DATE: ")[0], 250.), date_of_donation, font=content_font, fill=(255, 255, 255))

    draw.text((100,285), "RECIEPT_ID: ", font=content_font, fill=(255, 255, 255))
    draw.text((100 + content_font.getsize("RECIEPT_ID: ")[0], 285), str(reciept_id), font=content_font, fill=(255, 255, 255))


    # Set the starting position to print the receipt id


    # Print each line of the content on the certificate
    for line in content.split("\n"):
        if f"{count}" in line:
            # Use red color for count variable
            draw.text((x, y), line.split(f"{count}")[0], font=content_font, fill=(0, 0, 0))
            draw.text((x+content_font.getsize(line.split(f"{count}")[0])[0], y), f"{count}", font=content_font, fill=(255, 0, 0))
            draw.text((x+content_font.getsize(line.split(f"{count}")[0])[0]+content_font.getsize(f"{count}")[0], y), line.split(f"{count}")[1], font=content_font, fill=(0, 0, 0))

        elif f"{date_of_feeding}" in line:
            # Use blue color for date_of_feeding variable
            draw.text((x, y), line.split(f"{date_of_feeding}")[0], font=content_font, fill=(0, 0, 0))
            draw.text((x + content_font.getsize(line.split(f"{date_of_feeding}")[0])[0], y), f"{date_of_feeding}",
                      font=content_font, fill=(255, 0, 0))
            draw.text((x + content_font.getbbox(line.split(f"{date_of_feeding}")[0])[2] +
                       content_font.getbbox(f"{date_of_feeding}")[2], y), line.split(f"{date_of_feeding}")[1],
                      font=content_font, fill=(0, 0, 0))


        elif f"{price}" in line:
           # Use red color for price variable
            draw.text((x, y), line.split(f"{price}")[0], font=content_font, fill=(0, 0, 0))
            draw.text((x+content_font.getsize(line.split(f"{price}")[0])[0], y), f"{price}", font=content_font, fill=(255, 0, 0))
            draw.text((x+content_font.getsize(line.split(f"{price}")[0])[0]+content_font.getsize(f"{price}")[0], y), line.split(f"{price}")[1], font=content_font, fill=(0, 0, 0))


        else:
            draw.text((x, y), line, font=content_font, fill=(0, 0, 0))
        y += content_font.getbbox(line)[3] + 15  # Add the line height and some extra padding



    # Save the generated certificate

    if not os.path.exists("certificates"):
        os.mkdir("certificates")

    img.save(f"certificates/{name}_certificate.jpg")
