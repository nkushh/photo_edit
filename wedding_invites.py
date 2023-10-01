import os, time
import openpyxl, pywhatkit
from os.path import isfile, join
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont

class WeddingInvites:
    def __init__(self):
        self.file_extensions = ('.jpg', '.png', '.jpeg')
        self.invite_card = 'Julie_n_Gee.jpg'
        self.invite_list = openpyxl.load_workbook('invites_list.xlsx')

    def sendwhatsapp_msg(self, msg, phone_no, image_name):
        message_sent = pywhatkit.sendwhats_image(phone_no, image_name, msg, 1, True)
        time.sleep(10)
        if message_sent:
            return

    def run(self):
        # Custom font style and font size
        myFont = ImageFont.truetype("cac_champagne.ttf", 70, encoding='utf-8')
        myFontTwo = ImageFont.truetype("GoldleafBoldPersonalUseBold-eZ4dO.ttf",40)

        sheet_obj = self.invite_list.active
        m_row = sheet_obj.max_row

        for i in range(1, m_row + 1):
            # Open an Image
            img = Image.open(self.invite_card)
            
            # Call draw Method to add 2D graphics in an image
            myImage = ImageDraw.Draw(img)
            # I2 = ImageDraw.Draw(img)

            cell_obj = sheet_obj.cell(row = i, column = 1)
            # no_of_guests = sheet_obj.cell(row = i, column = 2)
            phone_no = sheet_obj.cell(row = i, column = 2)

            guest_name = cell_obj.value
            # card_admission = str(no_of_guests.value)
            # admission_text = f"This card admits only {card_admission}"
            print(f"Guest: {guest_name}")
            img_width = img.width/2
            guest_name_size = myImage.textbbox((0,0), guest_name, font=myFont)
            position = (img_width, 930)
            centered_position = (position[0] - guest_name_size[2] / 2, position[1] - guest_name_size[3] / 2)
            msg = f"Hallo {guest_name}, you are cordially invited"
            # # Add Text to an image
            myImage.text(centered_position, guest_name, font=myFont, fill=(248, 233, 177), spacing=2)
            # # I2.text((500, 870), admission_text, font=myFontTwo, fill=(45, 46, 46), align='center')
            # # Save the edited image
            image_name = f"{guest_name}.jpg"
            img.save(image_name)
            self.sendwhatsapp_msg(msg, phone_no.value, image_name)

# def main():
#     pass

# if __name__ == '__main__':
#     main()

guest_invite = WeddingInvites()
guest_invite.run()