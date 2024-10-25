#!/usr/bin/env python
import nltk, os, subprocess, code, glob, re, traceback, sys, inspect
# from time import clock, sleep
from pprint import pprint
import json
import zipfile
import docx2txt
# import ner
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
# from cStringIO import StringIO
from io import StringIO, BytesIO
from docx import Document
#from convertRtfToText import convertRtfToText
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from flask import Flask,request
from flask_cors import CORS, cross_origin
import urllib.request

class exportToCSV:
    def __init__(self, fileName='resultsCSV.txt', resetFile=False):
        headers = ['FILE NAME',
               'NAME',
               'EMAIL1', 'EMAIL2', 'EMAIL3', 'EMAIL4',
               'PHONE1', 'PHONE2', 'PHONE3', 'PHONE4',
               'INSTITUTES1','YEARS1',
               'INSTITUTES2','YEARS2',
               'INSTITUTES3','YEARS3',
               'INSTITUTES4','YEARS4',
               'INSTITUTES5','YEARS5',
               'EXPERIENCE',
               'DEGREES',
               ]
        if not os.path.isfile(fileName) or resetFile:
            # Will create/reset the file as per the evaluation of above condition
            fOut = open(fileName, 'w')
            fOut.close()
        fIn = open(fileName) ########### Open file if file already present
        inString = fIn.read()
        fIn.close()
        if len(inString) <= 0: ######### If File already exsists but is empty, it adds the header
            fOut = open(fileName, 'w')
            fOut.write(','.join(headers)+'\n')
            fOut.close()

    def write(self, infoDict):
        fOut = open('resultsCSV.txt', 'a+')
        # Individual elements are dictionaries
        writeString = ''
        try:
            writeString += str(infoDict['fileName']) + ','
            writeString += str(infoDict['name']) + ','
            
            if infoDict['email']:
                writeString += str(','.join(infoDict['email'][:4])) + ','
            if len(infoDict['email']) < 4:
                writeString += ','*(4-len(infoDict['email']))
            if infoDict['phone']:
                writeString += str(','.join(infoDict['phone'][:4])) + ','
            if len(infoDict['phone']) < 4:
                writeString += ','*(4-len(infoDict['phone']))            
            writeString += str(infoDict['%sinstitute'%'c\\.?a'])+","
            writeString +=str(infoDict['%syear'%'c\\.?a'])+","
            writeString += str(infoDict['%sinstitute'%'b\\.?com'])+","
            writeString +=str(infoDict['%syear'%'b\\.?com'])+","
            writeString += str(infoDict['%sinstitute'%'icwa'])+","
            writeString +=str(infoDict['%syear'%'icwa'])+","
            writeString += str(infoDict['%sinstitute'%'m\\.?com'])+","
            writeString +=str(infoDict['%syear'%'m\\.?com'])+","
            writeString += str(infoDict['%sinstitute'%'mba'])+","
            writeString +=str(infoDict['%syear'%'mba'])+","
            writeString += str(infoDict['experience']) + ','
            writeString += str(infoDict['degree']) + '\n' # For the remaining elements
            fOut.write(writeString)
        except:
            fOut.write('FAILED_TO_WRITE\n')
        fOut.close()

class Parse():
    # List (of dictionaries) that will store all of the values
    # For processing purposes
    information=[]
    inputString = ''
    tokens = []
    lines = []
    sentences = []
    jsonData = []

    def __init__(self, jdPath, resumePaths):
        fields = ["name", "address", "email", "phone", "mobile", "telephone", "residence status","experience","degree","cainstitute","cayear","caline","b.cominstitute","b.comyear","b.comline","icwainstitue","icwayear","icwaline","m.cominstitute","m.comyear","m.comline","mbainstitute","mbayear","mbaline"]
        
        self.jdPath = jdPath
        self.resumePaths = resumePaths
        
        files = self.resumePaths
        self.jsonData = []
        job_description = self.readFile(self.jdPath)[0]
        # clean_jd = self.clean_files(job_description) # resume ranking score is less after cleaning
        # clean_jd = ' '.join(str(e) for e in clean_jd) # resume ranking score is less after cleaning

        for f in files:
            # info is a dictionary that stores all the data obtained from parsing
            info = {}
            
            self.inputString = self.readFile(f["resumeFileUrl"])[0]
            info['extension'] = self.readFile(f["resumeFileUrl"])[1]
            info['fileName'] = f["resumeFileUrl"]

            self.tokenize(self.inputString)

            self.getEmail(self.inputString, info)

            self.getPhone(self.inputString, info)

            self.getName(self.inputString, info)

            self.getExperience(self.inputString, info)
            # csv=exportToCSV()
            # csv.write(info)
            self.information.append(info)
            # print (info)
            
            # clean_resume = self.clean_files(self.inputString) # resume ranking score is less after cleaning
            # clean_resume = ' '.join(str(e) for e in clean_resume) # resume ranking score is less after cleaning
            # text = [clean_resume,clean_jd] # resume ranking score is less after cleaning
            text = [self.inputString,job_description] # resume ranking score is more after cleaning
    
            ## Get a Match
            resumeScore = self.get_resume_score(text)

            extractedData = {}
            extractedData["name"] = info["name"]
            extractedData["phoneNo"] = list(dict.fromkeys(info["phone"]))
            extractedData["email"] = list(dict.fromkeys(info["email"]))
            # extractedData["fileName"] = info["fileName"].split("\\")[1]
            extractedData["resumeFileUrl"] = info["fileName"]
            extractedData["resumeScore"] = resumeScore
            extractedData["resumeFileName"] = f["resumeFileName"]
            extractedData["jdFileUrl"] = self.jdPath
            extractedData['experience'] = info["experience"]
            self.jsonData.append(extractedData)

    def sendData(self):
        return self.jsonData
                  
    def readFile(self, fileName):
        '''
        Read a file given its name as a string.
        Modules required: os
        UNIX packages required: antiword, ps2ascii
        '''
        extension = fileName.split(".")[-1]

        if extension == "txt":
            f = open(fileName, 'r')
            string = f.read()
            f.close() 
            return string, extension
        elif extension == "doc":
            # Run a shell command and store the output as a string
            # Antiword is used for extracting data out of Word docs. Does not work with docx, pdf etc.
            return subprocess.Popen(['antiword', fileName], stdout=subprocess.PIPE, stderr=subprocess.PIPE).communicate()[0], extension
        elif extension == "docx":
            try:
                return self.read_word_resume_from_url(fileName), extension
            except:
                return ''
                pass
        #elif extension == "rtf":
        #    try:
        #        return convertRtfToText(fileName), extension
        #    except:
        #        return ''
        #        pass
        elif extension == "pdf":
            # ps2ascii converst pdf to ascii text
            # May have a potential formatting loss for unicode characters
            # return os.system(("ps2ascii %s") (fileName))
            try:
                return self.read_text_from_pdf_url(fileName), extension
            except:
                return ''
                pass
        else:
            print('Unsupported format')
            return '', ''

    def convertPDFToText(self,pdf_doc):
        resource_manager = PDFResourceManager()
        fake_file_handle = StringIO()
        converter = TextConverter(resource_manager, fake_file_handle)
        page_interpreter = PDFPageInterpreter(resource_manager, converter)
    
        with open(pdf_doc, 'rb') as fh:
            for page in PDFPage.get_pages(fh, 
                                        caching=True,
                                        check_extractable=True):
                page_interpreter.process_page(page)
            
            text = fake_file_handle.getvalue()
    
        # close open handles
        converter.close()
        fake_file_handle.close()
    
        if text:
            return text

    def preprocess(self, document):
        '''
        Information Extraction: Preprocess a document with the necessary POS tagging.
        Returns three lists, one with tokens, one with POS tagged lines, one with POS tagged sentences.
        Modules required: nltk
        '''
        try:
            # Try to get rid of special characters
            try:
                document = document.decode('ascii', 'ignore')
            except:
                document = document.encode('ascii', 'ignore')
            # document = str.encode(document)
            document = document.decode('ascii', 'ignore')
            # Newlines are one element of structure in the data
            # Helps limit the context and breaks up the data as is intended in resumes - i.e., into points
            lines = [el.strip() for el in document.split("\n") if len(el) > 0]  # Splitting on the basis of newlines 
            lines = [nltk.word_tokenize(el) for el in lines]    # Tokenize the individual lines
            lines = [nltk.pos_tag(el) for el in lines]  # Tag them
            # Below approach is slightly different because it splits sentences not just on the basis of newlines, but also full stops 
            # - (barring abbreviations etc.)
            # But it fails miserably at predicting names, so currently using it only for tokenization of the whole document
            sentences = nltk.sent_tokenize(document)    # Split/Tokenize into sentences (List of strings)
            sentences = [nltk.word_tokenize(sent) for sent in sentences]    # Split/Tokenize sentences into words (List of lists of strings)
            tokens = sentences
            sentences = [nltk.pos_tag(sent) for sent in sentences]    # Tag the tokens - list of lists of tuples - each tuple is (<word>, <tag>)
            # Next 4 lines convert tokens from a list of list of strings to a list of strings; basically stitches them together
            dummy = []
            for el in tokens:
                dummy += el
            tokens = dummy
            # tokens - words extracted from the doc, lines - split only based on newlines (may have more than one sentence)
            # sentences - split on the basis of rules of grammar
            return tokens, lines, sentences
        except Exception as e:
            print(e) 

    def tokenize(self, inputString):
        try:
            self.tokens = self.preprocess(inputString)[0]
            self.lines = self.preprocess(inputString)[1]
            self.sentences = self.preprocess(inputString)[2]
            return self.tokens, self.lines, self.sentences
        except Exception as e:
            print(e)

    def getEmail(self, inputString, infoDict, debug=False): 
        '''
        Given an input string, returns possible matches for emails. Uses regular expression based matching.
        Needs an input string, a dictionary where values are being stored, and an optional parameter for debugging.
        Modules required: clock from time, code.
        '''

        email = None
        try:
            pattern = re.compile(r'\S*@\S*')
            matches = pattern.findall(inputString) # Gets all email addresses as a list
            email = matches
        except Exception as e:
            print(e)

        infoDict['email'] = email

        if debug:
            print("\n"), pprint(infoDict), "\n"
            code.interact(local=locals())
        return email

    def getPhone(self, inputString, infoDict, debug=False):
        '''
        Given an input string, returns possible matches for phone numbers. Uses regular expression based matching.
        Needs an input string, a dictionary where values are being stored, and an optional parameter for debugging.
        Modules required: clock from time, code.
        '''

        number = None
        try:
            pattern = re.compile(r'([+(]?\d+[)\-]?[ \t\r\f\v]*[(]?\d{2,}[()\-]?[ \t\r\f\v]*\d{2,}[()\-]?[ \t\r\f\v]*\d*[ \t\r\f\v]*\d*[ \t\r\f\v]*)')
                # Understanding the above regex
                # +91 or (91) -> [+(]? \d+ -?
                # Metacharacters have to be escaped with \ outside of character classes; inside only hyphen has to be escaped
                # hyphen has to be escaped inside the character class if you're not incidication a range
                # General number formats are 123 456 7890 or 12345 67890 or 1234567890 or 123-456-7890, hence 3 or more digits
                # Amendment to above - some also have (0000) 00 00 00 kind of format
                # \s* is any whitespace character - careful, use [ \t\r\f\v]* instead since newlines are trouble
            match = pattern.findall(inputString)
            # match = [re.sub(r'\s', '', el) for el in match]
                # Get rid of random whitespaces - helps with getting rid of 6 digits or fewer (e.g. pin codes) strings
            # substitute the characters we don't want just for the purpose of checking
            match = [re.sub(r'[,.]', '', el) for el in match if len(re.sub(r'[()\-.,\s+]', '', el))>6]
                # Taking care of years, eg. 2001-2004 etc.
            match = [re.sub(r'\D$', '', el).strip() for el in match]
                # $ matches end of string. This takes care of random trailing non-digit characters. \D is non-digit characters
            match = [el for el in match if len(re.sub(r'\D','',el)) <= 15]
                # Remove number strings that are greater than 15 digits
            try:
                for el in list(match):
                    # Create a copy of the list since you're iterating over it
                    if len(el.split('-')) > 3: continue # Year format YYYY-MM-DD
                    for x in el.split("-"):
                        try:
                            # Error catching is necessary because of possibility of stray non-number characters
                            # if int(re.sub(r'\D', '', x.strip())) in range(1900, 2100):
                            if x.strip()[-4:].isdigit():
                                if int(x.strip()[-4:]) in range(1900, 2100):
                                    # Don't combine the two if statements to avoid a type conversion error
                                    match.remove(el)
                        except:
                            pass
            except:
                pass
            number = match
        except:
            pass

        infoDict['phone'] = number

        if debug:
            print("\n"), pprint(infoDict), "\n"
            code.interact(local=locals())
        return number

    def getName(self, inputString, infoDict, debug=False):
        '''
        Given an input string, returns possible matches for names. Uses regular expression based matching.
        Needs an input string, a dictionary where values are being stored, and an optional parameter for debugging.
        Modules required: clock from time, code.
        '''

        # Reads Indian Names from the file, reduce all to lower case for easy comparision [Name lists]
        indianNames = open("allNames.txt", "r").read().lower()
        # Lookup in a set is much faster
        indianNames = set(indianNames.split())

        otherNameHits = []
        nameHits = []
        name = None
        
        try:
            # tokens, lines, sentences = self.preprocess(inputString)
            tokens, lines, sentences = self.tokens, self.lines, self.sentences

            # Try a regex chunk parser
            # grammar = r'NAME: {<NN.*><NN.*>|<NN.*><NN.*><NN.*>}'
            grammar = r'NAME: {<NN.*><NN.*>}'
            # Noun phrase chunk is made out of two or three tags of type NN. (ie NN, NNP etc.) - typical of a name. {2,3} won't work, hence the syntax
            # Note the correction to the rule. Change has been made later.
            chunkParser = nltk.RegexpParser(grammar)
            all_chunked_tokens = []
            for tagged_tokens in lines:
                # Creates a parse tree
                if len(tagged_tokens) == 0: continue # Prevent it from printing warnings
                chunked_tokens = chunkParser.parse(tagged_tokens)
                all_chunked_tokens.append(chunked_tokens)
                for subtree in chunked_tokens.subtrees():
                    #  or subtree.label() == 'S' include in if condition if required
                    if subtree.label() == 'NAME':
                        for ind, leaf in enumerate(subtree.leaves()):
                            if leaf[0].lower() in indianNames and 'NN' in leaf[1]:
                                # Case insensitive matching, as indianNames have names in lowercase
                                # Take only noun-tagged tokens
                                # Surname is not in the name list, hence if match is achieved add all noun-type tokens
                                # Pick upto 3 noun entities
                                hit = " ".join([el[0] for el in subtree.leaves()[ind:ind+3]])
                                # Check for the presence of commas, colons, digits - usually markers of non-named entities 
                                if re.compile(r'[\d,:]').search(hit): continue
                                nameHits.append(hit)
                                # Need to iterate through rest of the leaves because of possible mis-matches
            # Going for the first name hit
            if len(nameHits) > 0:
                nameHits = [re.sub(r'[^a-zA-Z \-]', '', el).strip() for el in nameHits] 
                name = " ".join([el[0].upper()+el[1:].lower() for el in nameHits[0].split() if len(el)>0])
                otherNameHits = nameHits[1:]

        except Exception as e:
            print(traceback.format_exc())
            print(e)        

        infoDict['name'] = name
        infoDict['otherNameHits'] = otherNameHits

        if debug:
            print("\n"), pprint(infoDict), "\n"
            code.interact(local=locals())
        return name, otherNameHits  
    
    def getExperience(self,inputString,infoDict,debug=False):
        pattern = r'(\d+)\s+years?\s+(\d+)\s+months?'

        # Search for the pattern in the resume text
        match = re.search(pattern, inputString)

        if match:
            years = int(match.group(1))
            months = int(match.group(2))
            total_experience = years + (months / 12)
            infoDict['experience'] = round(total_experience,2)
            return round(total_experience,2)
        else:
            infoDict['experience'] = "Experience not found in the resume text."
            return "Experience not found in the resume text."

    def get_resume_score(self,text):
        # cv = CountVectorizer(stop_words='english') # score is less using stopwords
        cv = CountVectorizer() # score is more without using stopwords
        count_matrix = cv.fit_transform(text)
        #get the match percentage
        matchPercentage = cosine_similarity(count_matrix)[0][1] * 100
        matchPercentage = round(matchPercentage, 2) # round to two decimal
        return matchPercentage
    
    def clean_files(self,jd):
        ''' a function to create a word cloud based on the input text parameter'''
        ## Clean the Text
        # Lower

        clean_jd = jd.lower()
        # remove punctuation
        clean_jd = re.sub(r'[^\w\s]', '', clean_jd)
        # remove trailing spaces
        clean_jd = clean_jd.strip()
        # remove numbers
        clean_jd = re.sub('[0-9]+', '', clean_jd)
        # tokenize 
        clean_jd = word_tokenize(clean_jd)
        # remove stop words
        stop = stopwords.words('english')
        #stop.extend(["AT_USER","URL","rt","corona","coronavirus","covid","amp","new","th","along","icai","would","today","asks"])
        clean_jd = [w for w in clean_jd if not w in stop] 
    
        return(clean_jd)
    
    def convertDocxToText(self,path):
        document = Document(path)
        return "\n".join([para.text for para in document.paragraphs])

    def read_text_from_pdf_url(self,url, user_agent=None):
        resource_manager = PDFResourceManager()
        fake_file_handle = StringIO()
        converter = TextConverter(resource_manager, fake_file_handle)    

        if user_agent == None:
            user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36'

        headers = {'User-Agent': user_agent}
        request = urllib.request.Request(url, data=None, headers=headers)

        response = urllib.request.urlopen(request).read()
        fb = BytesIO(response)

        page_interpreter = PDFPageInterpreter(resource_manager, converter)

        for page in PDFPage.get_pages(fb,
                                    caching=True,
                                    check_extractable=True):
            page_interpreter.process_page(page)


        text = fake_file_handle.getvalue()

        # close open handles
        fb.close()
        converter.close()   
        fake_file_handle.close()

        if text:
            # If document has instances of \xa0 replace them with spaces.
            # NOTE: \xa0 is non-breaking space in Latin1 (ISO 8859-1) & chr(160)
            text = text.replace(u'\xa0', u' ')
            return text

    def read_word_resume_from_url(self,url):
        #resume = docx2txt.process(word_doc)
        #text = ''.join(resume)
        word_doc = BytesIO(urllib.request.urlopen(url).read())
        resume = docx2txt.process(word_doc)
        resume = str(resume)
        #print(resume)
        text =  ''.join(resume)
        text = text.replace("\n", "")
        if text:
            return text

app = Flask(__name__)
CORS(app,supports_credentials=True,resources={r"/api/*": {"origins": "*"}})

@app.route('/resumescreening/screenResume',methods = ['POST'])
@cross_origin(supports_credentials=True,resources={r"/api/*": {"origins": "*"}})

def processData():
    data = request.get_json()
    jdPath = data['jdPath']
    resumePaths = data['resumePath']
    p = Parse(jdPath,resumePaths)
    jsonData = p.sendData()
    response = {
        "response" : jsonData
    }
    return response

if __name__ == '__main__':
    app.run(debug = True)