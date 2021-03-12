from PIL import Image
import xlsxwriter as excel
import time

def start_conversion(width, height, pixels):
    # Create a new Excel book
    book = excel.Workbook("output.xls")

    # Create a new worksheet inside the Excel book
    sheet1 = book.add_worksheet()

    # Set the sheet's cell size to 40 x 40
    sheet1.set_default_row(40)

    # Generate the Excel cells from the image
    letters = encode_numbers(width)
    for x in range(width):
        column_letters = get_letters_as_string(next(letters))
        for y in range(height):
            color_hex = rgb_to_hex(pixels[x, y])

            cell_format = book.add_format()
            cell_format.set_pattern(1)
            cell_format.set_bg_color(color_hex)

            str_encode = column_letters + str(y)
            sheet1.write(str_encode, "", cell_format)
            
    # Save and close the Excel book
    book.close()

# Generator function to encode number to letter index in Excel
def encode_numbers(num):
    if (num < 0):
        print("Invalid number to encode")
        exit(1)

    start = ['A', None, None]

    for i in range(num):
        if start[0] == 'Z':
            if start[1] == None:
                start = ['A', 'A', start[2]]
            else:
                if start[1] == 'Z':
                    if start[2] == None:
                        start = ['A', 'A', 'A']
                    else:
                        start = ['A', 'A', chr(ord(start[2]) + 1)]
                else:
                    start = ['A', chr(ord(start[1]) + 1), start[2]]
        else:
            start = [chr(ord(start[0]) + 1), start[1], start[2]]

        yield start
            

# Convert list of letters to string
def get_letters_as_string(start):
    start = start[::-1]
    if start[0] == None:
        if start[1] == None:
            return start[2]
        else:
            return start[1] + start[2]
    else:
        return ''.join(start)


# Convert (r, g, b) to hex
def rgb_to_hex(rgb):
    return '#%02x%02x%02x' % (rgb[0], rgb[1], rgb[2])


if __name__ == '__main__':
    # Get the file name from user
    file = input("Enter photo name: ")

    # Try and open the image
    im = None
    try:
        im = Image.open(file)
    except:
        print("Image not found")
        exit(1)

    # Get image width and height
    width, height = im.size
    print("Width={}, Height={}".format(width, height))

    # Let user know it's working
    print("Converting your image to Excel cells...")

    # Start timer 
    start_time = time.time()

    # Convert image to Excel cells
    start_conversion(width, height, im.load())

    # Print out how long it took
    end_time = time.time()
    print("Time taken: {}".format(end_time - start_time))
