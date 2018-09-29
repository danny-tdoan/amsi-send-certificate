'''
This script renames the PDF certificates and associates the files with the corresponding recipients.
'''
import sys
import os
import csv


def rename_pdf_files(cert_dir, filename_prefix, recipient_info):
    """Read the recipient infos, map the auto-generated PDF certificates and rename it to the
    format <Firstname><Lastname>.pdf

    This function also consolidate the information into a central csv that will be used to compose the email
    in the next step

    Input:
        1. Directory to all certificates
        2. Prefix of each filename (needs to be the same)
        3. recipient_info is a dictionary in the form {id1:{'firstname':'John','lastname':'Doe',...}}
    """


    # this is the final input used to send emails
    compiled_recipients = []

    # in case two people have the same first name and last name
    seen_names = []

    for id, r_info in recipient_info.items():
        try:
            new_filename = r_info['firstname'] + r_info['lastname']

            cnt = 1
            while new_filename in seen_names:
                print(
                    'A certificate for ' + new_filename + ' already exists. Probably duplicated names of participants. Appending number ' + str(
                        cnt) + ' to it.')
                new_filename = new_filename + str(cnt)
                cnt += 1

            seen_names.append(new_filename)

            new_filename += '.pdf'

            # rename the file
            new_filename_with_path = cert_dir + '/' + new_filename
            os.rename(cert_dir + '/' + filename_prefix + str(id) + '.pdf', new_filename_with_path)
        # print("Renamed "+cert_dir+'/'+filename_prefix+str(id)+'.pdf to '+ new_filename_with_path)
        except Exception as e:
            print(e)
            print("WARNING: Something wrong with the record for " + r_info['firstname'] + " " + r_info[
                'lastname'] + ". ID is " + str(id))
            print("Double check if the file is renamed.")

        try:
            compiled_recipients.append(
                [id, r_info['firstname'], r_info['lastname'], r_info['email'], new_filename_with_path])
        except Exception as e:
            print(e)
            print("WARNING: Something wrong when appending the recipient" + r_info['firstname'] + " " + r_info[
                'lastname'] + " to the list.")
            print("Check if the record is appended correctly.")

    # all done, generate the recipients csv file
    print("Writing the recipient infos to file. Ready to send emails.")
    with open('recipients.csv', 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(['id', 'firstname', 'lastname', 'email', 'attachment'])
        csvwriter.writerows(compiled_recipients)
    print("Recipients written to " + os.getcwd() + '\\recipients.csv')


def get_recipient_info(recipient_file):
    '''Extract the recipient info from the spreadsheet. Info only contains firstname, lastname, email'''

    recipient_info = {}
    with open(recipient_file) as csvfile:
        reader = csv.DictReader(csvfile)
        id = 1
        for row in reader:
            try:
                recipient_info[id] = {'firstname': row['First_Name'], 'lastname': row['Last_Name'],
                                      'email': row['Email']}
                id += 1
            except Exception as e:
                print(e)
                print(
                    'WARNING: Make sure the header of the csv file has these fields First_Name, Last_Name, Email (case sensitive)')

    return recipient_info


if __name__ == "__main__":
    cert_dir = sys.argv[1]
    recipient_list = sys.argv[2]

    print("The next step asks for file prefix. If the certificates are name BIS17_Certificates 1.pdf, "
          "BIS17_Certificates 2.pdf, BIS17_Certificates 3.pdf, etc. the prefix is 'BIS17_Certificates '"
          "(with the space at the end, and without quotes)")
    file_prefix = input("Enter file prefix: ")

    recipient_info = get_recipient_info(recipient_list)
    rename_pdf_files(cert_dir, file_prefix, recipient_info)
