import config
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import glob
import logging
from logging.handlers import RotatingFileHandler
import os
import pandas as pd
import pandas.io.formats.excel
import smtplib
import ssl
import sys
from time import localtime, strftime


def email_report(report_date):
    """
    This does the final formatting of the email and sends it
    """
    logging.info("Creating email...")

    port = 587
    smtp_server = "smtp.office365.com"
    login_email = config.email_login
    from_address = "it@bedrocklogistics.com"
    to_address = ["derrick.freeman@bedrocklogistics.com"]
    password = config.email_password

    # gets the most recently created file to the CurrentReport Folder (there should never be more than 1 file though)
    file_location = max(glob.glob("C:\\export\\data\\ftp\\CurrentReport\\*.*"), key=os.path.getctime)
    subject_and_body = "SPW Report " + report_date

    # create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = from_address
    message["To"] = ", ".join(to_address)
    message["Subject"] = subject_and_body

    # add body to the email
    message.attach(MIMEText(subject_and_body, "plain"))

    # Setup the attachment
    filename = os.path.basename(file_location)
    attachment = open(file_location, "rb")
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {filename}")

    # Attach the attachment to the MIMEMultipart object
    message.attach(part)

    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_server, port) as server:
        try:
            logging.info("Attempting to send email with report...")
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            server.login(login_email, password)
            server.sendmail(from_address, to_address, message.as_string())
            logging.info("Email has successfully sent.")
        except Exception as e:
            logging.exception(e)
        finally:
            server.quit()
            logging.info("Successfully quit server.")


def empty_directory():
    logging.info("Clearing CurrentReport folder to prepare for new file...")
    # Empty directory before saving current file
    try:
        currentreport_dir = glob.glob("C:\\export\\data\\ftp\\CurrentReport\\*")
        for file in currentreport_dir:
            os.remove(file)
        logging.info("CurrentReport folder is successfully emptied")
    except Exception:
        logging.exception("Unable to empty CurrentReport folder...")
        logging.info("CurrentReport folder does not need to be empty so script is continuing")


def log_setup():
    """
    Just for a easy reminder so the file doesn't get too large over time.
    """
    logging.basicConfig(
        handlers=[
            RotatingFileHandler(
                "./spw_load_export.log", mode="a", maxBytes=1 * 1024 * 1024, backupCount=1, encoding="utf-8", delay=0
            )
        ],
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        level=logging.INFO,
    )


def revise_file():
    logging.info("Revising file...")
    # read current days file
    df = pd.read_excel(max(glob.glob("C:\\export\\data\\ftp\\MercuryGate\\*.*"), key=os.path.getctime))
    filename = os.path.basename(max(glob.glob("C:\\export\\data\\ftp\\MercuryGate\\*.*"), key=os.path.getctime))

    # rename column name SKIP to #SKIP since Talend doesn't allow # in column names
    df.rename(columns={"SKIP": "#SKIP"}, inplace=True)

    # if mbl_pri_ref (column C) has an emtpy value fill it will the value from mbl_addl_ref (column E)
    if df["mbl_pri_ref"].isnull:
        df["mbl_pri_ref"] = df["mbl_pri_ref"].fillna(df["mbl_addl_ref"])

    # if shipment_pri_ref (column L) has an emtpy value fill it will the value from mbl_addl_ref (column E)
    if df["shipment_pri_ref"].isnull:
        df["shipment_pri_ref"] = df["shipment_pri_ref"].fillna(df["mbl_addl_ref"])

    """
        - This formats columns O - R (ship_start_date, ship_end_date, delivery_start_date, delivery_end_date) and keeps
        the columns without any dates empty

        - "%-m/%-d/%Y" linux or "%#m/%#d/%Y" for windows
    """
    df["ship_start_date"] = [d.strftime("%#m/%#d/%Y") if not pd.isnull(d) else "" for d in df["ship_start_date"]]
    df["ship_end_date"] = [d.strftime("%#m/%#d/%Y") if not pd.isnull(d) else "" for d in df["ship_end_date"]]
    df["delivery_start_date"] = [
        d.strftime("%#m/%#d/%Y") if not pd.isnull(d) else "" for d in df["delivery_start_date"]
    ]
    df["delivery_end_date"] = [d.strftime("%#m/%#d/%Y") if not pd.isnull(d) else "" for d in df["delivery_end_date"]]

    # get the date the report was for which is the ship_start_date
    report_date = df.iloc[0]["ship_start_date"]

    # replace na with nothing like it is originally
    df = df.fillna("")

    # convert all columns to strings
    all_columns = list(df)
    df[all_columns] = df[all_columns].astype(str)

    # removes the default formatting done by pandas (e.g. the header being bold and outlined)
    pandas.io.formats.excel.ExcelFormatter.header_style = None

    logging.info("File revisions done...")
    logging.info("Saving revised file to CurrentReport folder...")

    # save report so we can attach to email later
    try:
        df.to_excel(f"C:\\export\\data\\ftp\\CurrentReport\\{filename}", index=False)
        logging.info("File saved...")
    except Exception:
        logging.exception("Unable to save file...")
        sys.exit("Unable to save file, ending script...")

    email_report(report_date)


def talend_job():
    # see if the cron job that runs the job ran
    most_recent_file_date = strftime(
        "%m/%d/%Y",
        localtime(os.path.getctime(max(glob.glob("C:\\export\\data\\ftp\\MercuryGate\\*.*"), key=os.path.getctime))),
    )
    today = strftime("%m/%d/%Y", localtime())

    if most_recent_file_date == today:
        logging.info("spw_load_export.xls cron job has ran and created a new file today.")
        logging.info("Continuing script...")
    else:
        logging.critical("spw_load_export.xls talend job has not ran and/or created a new file. Unable to continue.")
        sys.exit("New file not found, ending script...")


def main():
    # Empty dir before saving current finished file. This dir holds report for sending without replacing unedited file
    empty_directory()

    # see if talend job was run
    talend_job()

    # revise excel file
    revise_file()


if __name__ == "__main__":
    log_setup()
    logging.info("spw_load_export.py script has started.")
    main()
