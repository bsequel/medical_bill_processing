import sys
import os
from urllib.parse import urlparse
import boto3
import time
from aws_lib_.tdp import DocumentProcessor
from aws_lib_.og import OutputGenerator
from aws_lib_.helper import FileHelper, S3Helper
import boto3
from botocore.config import Config
import boto3
import shutil
import pandas as pd
import time
import re
import shutil
import os
import math
from tika import parser
import openpyxl
import psycopg2
import datetime
import codecs


class Textractor:
    def getInputParameters(self, args):
        event = {}
        i = 0
        print(args)
        args=args.split()            
        if(args):
            while(i < len(args)):
                if(args[i] == '--documents'):
                    event['documents'] = args[i+1]
                    i = i + 1
                if(args[i] == '--region'):
                    event['region'] = args[i+1]
                    i = i + 1
                if(args[i] == '--text'):
                    event['text'] = True
                if(args[i] == '--forms'):
                    event['forms'] = True
                if(args[i] == '--tables'):
                    event['tables'] = True
                if(args[i] == '--insights'):
                    event['insights'] = True
                if(args[i] == '--medical-insights'):
                    event['medical-insights'] = True
                if(args[i] == '--translate'):
                    event['translate'] = args[i+1]
                    i = i + 1
                i = i + 1
        print(event)
        return event

    def validateInput(self, args):
        event = self.getInputParameters(args)

        ips = {}

        if(not 'documents' in event):
            raise Exception("Document or path to a foler or S3 bucket containing documents is required.")

        inputDocument = event['documents']
        idl = inputDocument.lower()
        bucketName = None
        documents = []
        awsRegion = 'us-east-1'

        if(idl.startswith("s3://")):
            o = urlparse(inputDocument)
            bucketName = o.netloc
            path = o.path[1:]
            ar = S3Helper.getS3BucketRegion(bucketName)
            if(ar):
                awsRegion = ar

            if(idl.endswith("/")):
                allowedFileTypes = ["jpg", "jpeg", "png", "pdf"]
                documents = S3Helper.getFileNames(awsRegion, bucketName, path, 1, allowedFileTypes)
            else:
                documents.append(path)
        else:
            if(idl.endswith("/")):
                allowedFileTypes = ["jpg", "jpeg", "png"]
                documents = FileHelper.getFileNames(inputDocument, allowedFileTypes)
            else:
                documents.append(inputDocument)

            if('region' in event):
                awsRegion = event['region']

        ips["bucketName"] = bucketName
        ips["documents"] = documents
        ips["awsRegion"] = awsRegion
        ips["text"] = ('text' in event)
        ips["forms"] = ('forms' in event)
        ips["tables"] = ('tables' in event)
        ips["insights"] = ('insights' in event)
        ips["medical-insights"] = ('medical-insights' in event)
        if("translate" in event):
            ips["translate"] = event["translate"]
        else:
            ips["translate"] = ""

        return ips

    def processDocument(self, ips, i, document):
        print("\nTextracting Document # {}: {}".format(i, document))
        print('=' * (len(document)+30))

        # Get document textracted
        dp = DocumentProcessor(ips["bucketName"], document, ips["awsRegion"], ips["text"], ips["forms"], ips["tables"])
        response = dp.run()

        if(response):
            print("Recieved Textract response...")
            #FileHelper.writeToFile("temp-response.json", json.dumps(response))

            #Generate output files
            print("Generating output...")
            name, ext = FileHelper.getFileNameAndExtension(document)
            opg = OutputGenerator(response,
                        "{}-{}".format(name, ext),
                        ips["forms"], ips["tables"])
            opg.run()

            if(ips["insights"] or ips["medical-insights"] or ips["translate"]):
                opg.generateInsights(ips["insights"], ips["medical-insights"], ips["translate"], ips["awsRegion"])

            print("{} textracted successfully.".format(document))
        else:
            print("Could not generate output for {}.".format(document))

    def printFormatException(self, e):
        print("Invalid input: {}".format(e))
        print("Valid format:")
        print('- python3 textractor.py --documents mydoc.jpg --text --forms --tables --region us-east-1')
        print('- python3 textractor.py --documents ./myfolder/ --text --forms --tables')
        print('- python3 textractor.py --documents s3://mybucket/mydoc.pdf --text --forms --tables')
        print('- python3 textractor.py --documents s3://mybucket/ --text --forms --tables')

    def run(self,argv):

        ips = None
        try:
            ips = self.validateInput(argv)
        except Exception as e:
            self.printFormatException(e)

        #try:
        i = 1
        totalDocuments = len(ips["documents"])

        print("\n")
        print('*' * 60)
        print("Total input documents: {}".format(totalDocuments))
        print('*' * 60)

        for document in ips["documents"]:
        
            self.processDocument(ips, i, document)
     

            remaining = len(ips["documents"])-i

            if(remaining > 0):
                print("\nRemaining documents: {}".format(remaining))

                # print("\nTaking a short break...")
                # time.sleep(20)
                # print("Allright, ready to go...\n")

            i = i + 1

        print("\n")
        print('*' * 60)
        print("Successfully textracted documents: {}".format(totalDocuments))
        print('*' * 60)
        print("\n")
        #except Exception as e:
        #    print("Something went wrong:\n====================================================\n{}".format(e))

#Textractor().run()
def upload_cloud(pdf,name):
    s3 = boto3.resource('s3')
    BUCKET = "sequelbucket2022"
    s3.Bucket(BUCKET).upload_file(pdf, name)
    print("uploaded")

def api_request(name):
    s=Textractor()
    #s.getInputParameters(r'--documents s3://testpdfs1200/AKNA__PO_2112_GGN47.pdf --text --forms --tables')
    print(f'--documents s3://sequelbucket2022/{name}')
    print(f'--documents s3://sequelbucket2022/{name} --text')
    s.validateInput(f'--documents s3://sequelbucket2022//{name} --text --tables')
    s.run(f'--documents s3://sequelbucket2022/{name} --text --tables')
    print("done")

def main_call(input_pdf_path):

   if  '.jpeg' in input_pdf_path or '.pdf' in input_pdf_path.lower() or '.PDF' in input_pdf_path or '.png' in input_pdf_path or '.JPEG' in input_pdf_path or '.jpg' in input_pdf_path.lower()  or '.tif' in input_pdf_path.lower()  or '.tiff' in input_pdf_path.lower() : 
       pdf=input_pdf_path
       name=input_pdf_path.split("\\")[-1]
       pdf= input_pdf_path
       upload_cloud(pdf,name)
       print("uploaded")
       api_request(name)




