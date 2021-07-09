#!/usr/bin/python
# -*- coding: utf-8 -*-
import re
import csv
import os
from os import listdir
from os.path import isfile, join
import sys;
import getopt
import zipfile
from xml.dom import minidom
import logging

debuglevel= "info"

def print_help():
    print("Help:")
    print("--help or -h: print this information");
    print("--input=docxfile.docx or -i docxfile.docx: give file on which to operate)");
    print("--dirinput=dirname or -d dirname: scan a whole directory of docx files")
    print("--verbose or -v: tell me everything about what you are doing")
    print("--scan or -s: just look through the file(s) and report on presence or absence of docx bookmarks")
    print("--remove or -r: remove those bookmarklets and try to replace them with the currently selected texts")

def seek_calcfields(thefilename):
    global logging; 
    calcfields=0;
    listentry=0;
    fields=False;
    zip = zipfile.ZipFile(thefilename);
    #print (zip.namelist())
    logging.debug(str(zip.namelist()));
    # extract the word/document.xml 
    f = zip.open("word/document.xml")
    content = f.read()
    mydoc = minidom.parseString(content)
    
    pretty_xml_as_string = mydoc.toprettyxml()
    #print(pretty_xml_as_string);
    for value in mydoc.getElementsByTagName("w:fldChar"):
        calcfields=calcfields+1;
        fields=True;
    for value in mydoc.getElementsByTagName("w:listEntry"):
        listentry=listentry+1;
        fields=True;
    #test=xml.dom.minidom.parseString(content).toprettyxml()
    #print(test);
    return(calcfields, fields,listentry);

def seek_bookmarks(thefilename):
    bookmarkcount=0;
    bookmarks=False;
    zip = zipfile.ZipFile(thefilename);
    print (zip.namelist())
    # extract the word/document.xml 
    f = zip.open("word/document.xml")
    content = f.read()
    mydoc = minidom.parseString(content)
    
    pretty_xml_as_string = mydoc.toprettyxml()
    for value in mydoc.getElementsByTagName("w:bookmarkStart"):
        bookmarkcount=bookmarkcount+1;
        bookmarks=True;
    return(bookmarks, bookmarkcount);

def remove_calcfields(thefilename):
    logging.debug(print("Working on "+thefilename));
    #copyfile(thefilename,thefilename+"-nocalc.docx");
    modifiedfilename=thefilename+"-nocalc.docx";
    zip=zipfile.ZipFile(thefilename);
    logging.debug("Files in docx: "+str(zip.namelist()));
    f=zip.open("word/document.xml");
    content=f.read();
    mydoc=minidom.parseString(content);
    #logging.debug("Docx content: "+str(mydoc.toprettyxml())); 
    #test=mydoc.toprettyxml()

    # Also, you get text in strikethroughs. TODO for future work, provide mode to remove all text in strikethroughs... 
    # these are corrections... 

    # there is also a text element case that may appear 
    # <w:r w:rsidRPr="009B7433">
    #    <w:rPr>
    #      <w:rFonts w:ascii="Book Antiqua" w:hAnsi="Book Antiqua" w:cs="Arial"/>
    #      <w:b/>
    #    </w:rPr>
    #    <w:fldChar w:fldCharType="begin">
    #      <w:ffData>
    #        <w:name w:val="Text7"/>
    #        <w:enabled/>
    #        <w:calcOnExit w:val="0"/>
    #        <w:textInput>
    #          <w:default w:val="THE SECRETARY OF STATE FOR THE HOME DEPARTMENT"/>
    #          <w:format w:val="UPPERCASE"/>
    #        </w:textInput>
    #      </w:ffData>
    #    </w:fldChar>
    #  </w:r>
    # node should be replaced with a node that just contains the default val of the text input, probably
    #for par in mydoc.getElementsByTagName("w:p"): # for each paragraph
    #    if(par.getElementsByTagName("w:textInput")):
    #        print(par.toprettyxml());
    #        print("textinput found");
    #        textInputs=par.getElementsByTagName("w:textInput");
    #        for ti in textInputs:
    #            ffd=ti.parentNode;
    #            print ffd.toprettyxml();
    # currently I haven't handled this case because it seems to be readable unaltered by text extraction software 

    # let's deal with dropdown menus 
    for par in mydoc.getElementsByTagName("w:p"): # for each paragraph
        if(par.getElementsByTagName('w:fldChar')):
            logging.debug(str(par.toprettyxml()));
        # for each fldChar, which is the field that contains the data selector: 
        for fldchars in par.getElementsByTagName('w:fldChar'):
            responseindex=0;
            selectedEntry=""
            if(fldchars.getElementsByTagName('w:ffData')):
                logging.debug(str(fldchars.toprettyxml()));
                ffData=fldchars.getElementsByTagName('w:ffData');
                
            if(fldchars.getElementsByTagName('w:ddList')): # only under these circumstances can we do anything with it afaik
                logging.debug(str("Contains ddList"));
                ddList=fldchars.getElementsByTagName('w:ddList');
                # this can contain a w:result and it can contain listEntry elements. It would appear that if there is a w:result with a numeric w:val then that is the zero-indexed node from listEntry to choose. 
                candidateValues=[];
                for dd in ddList:
                    resultval=dd.getElementsByTagName('w:result'); # may not exist if not calculated, i.e. nobody actually ever used the pulldown menu. If not present then the item that is showing will be the top item in the list *shrug* so select this one
                    if(len(resultval)>0):
                        responseindex=resultval[0].getAttribute("w:val");
                    listEntries=dd.getElementsByTagName('w:listEntry');
                    logging.debug(str("Retrieving list item "+str(responseindex)));
                    for le in listEntries:
                        logging.debug(str(le.toprettyxml()));      
                        candidateValues.append(le.getAttribute('w:val'));
                    selectedEntry=candidateValues[int(responseindex)];
            logging.debug(str("Selected entry is : "+str(selectedEntry)));
            if(len(selectedEntry)>0):
                # Get parent node of fldChars
                wr=fldchars.parentNode;
                wrp=wr.parentNode;
                # Also don't forget that you need to copy the paragraph context from the one you're removing to the one you're adding 
                parcontext=wr.getAttribute("w:rsidRPr")
                newNode=mydoc.createElement("w:r")
                newNode.setAttribute("w:rsidRPr",parcontext);
                newRprNode=mydoc.createElement("w:rPr")
                newFontElement=mydoc.createElement("w:rFonts")
                # should probably steal the font info from another node, but really I suppose it hardly matters for our purposes (which, for anyone reading this, is just to support text extraction)
                newFontElement.setAttribute("w:ascii","Sylfaen");
                newFontElement.setAttribute("w:hAnsi","Sylfaen");
                newFontElement.setAttribute("w:cs","Arial");
                newNode.appendChild(newRprNode);
                newRprNode.appendChild(newFontElement);
                newWtNode=mydoc.createElement("w:t")
                newWtNode.setAttribute("xml:space","preserve")
                newTextNode=mydoc.createTextNode(" "+selectedEntry+" ");
                newWtNode.appendChild(newTextNode);
                newNode.appendChild(newWtNode);
                logging.debug("Candidate node")
                logging.debug(newNode.toprettyxml());
                wrp.insertBefore(newNode,wr)
                wrp.removeChild(wr) 
                if(par.getElementsByTagName('w:instrText')):
                     for instrT in par.getElementsByTagName('w:instrText'):
                        wr=instrT.parentNode;
                        wrp=wr.parentNode;
                        wrp.removeChild(wr);


    # NOW that we are done with our replacements, create a zipfile that contains the modified version: 
    with zipfile.ZipFile(thefilename) as inzip, zipfile.ZipFile(modifiedfilename, "w") as outzip:
        for zip_info in inzip.infolist():
            if zip_info.filename=="word/document.xml":
                outzip.writestr(zip_info.filename,mydoc.toxml().encode("utf-8"));
            else: 
                with inzip.open(zip_info.filename) as infile:
                    content = infile.read()
                    outzip.writestr(zip_info.filename, content)

def scan_directory(thedirname):
    scan_results=[];
    onlyfiles = [f for f in listdir(thedirname) if isfile(join(thedirname, f)) ]
    for f in onlyfiles:
        lowerf=f.lower();
        if lowerf.endswith('.docx'):
            scan_results.append(os.path.join(thedirname,f))
    return scan_results;

if __name__ == "__main__":
    # Get command line options 
    #logging.basicConfig(filename='example.log', encoding='utf-8', level=logging.DEBUG)
    logging.basicConfig(filename='sample.log', level=logging.INFO)
    arguments = len(sys.argv) - 1
    full_cmd_arguments = sys.argv
    argument_list = full_cmd_arguments[1:]
    short_opts="hi:d:vsr" 
    long_opts=["help","input=","dirinput=","verbose","scan","remove"]; 

    try:
        arguments, values = getopt.getopt(argument_list, short_opts, long_opts)
        #print(arguments);
    except getopt.error as err:
        # Output error, and return with an error code
        print (str(err))
        sys.exit(2)
    
    target_files=[];
    for current_argument, current_val in arguments:
         
        logging.debug(current_argument+", "+current_val)
        if(current_argument=="--verbose" or current_argument=="-v"):
            debuglevel="debug";
            logging.basicConfig(filename='sample.log', level=logging.DEBUG)
        if(current_argument=="--help" or current_argument=="-h"):
            print_help();
            sys.exit();

    for current_argument, current_val in arguments:
        if(current_argument=="--input" or current_argument=="-i"):
            target_files.append(current_val);
            logging.debug(str(target_files));
        if(current_argument=="--dirinput" or current_argument=="-d"):
            # scan through directory current_val to look for docx files, add all to list
            if(not os.path.isdir(current_val)):
                print("Object to scan must be directory")
                sys.exit()
            else:
                print("Scanning")
                target_files=scan_directory(current_val);
            
    print(target_files)
    for current_argument, current_val in arguments:
        if(current_argument=="--scan" or current_argument=="-s"):
            total_files=0;
            total_files_with_calcfields=0;
            total_files_with_listentries=0;
            calcfields_per_file=[];
            listentries_per_file=[];
            
            print("Scan");
            logging.debug(str(target_files));
            for target_file in target_files:
                # Scanning for calculated fields. 
                (calcfields,calcfieldscount,listentriescount)=seek_calcfields(target_file);
                #print(calcfields)
                #print(calcfieldscount);
                if(calcfields):
                    total_files_with_calcfields=total_files_with_calcfields+1;
                if(listentriescount>0):
                    total_files_with_listentries=total_files_with_listentries+1;
                calcfields_per_file.append(calcfieldscount);
                listentries_per_file.append(listentriescount);
                total_files=total_files+1;

            # now calculate and present stats
            print("Of a total of "+str(total_files)+" files, "+str(total_files_with_calcfields) + " contained calcfields, of which "+str(total_files_with_listentries)+" contain list selectors");
            # generate a report: which specific filenames contain calcfields?
            print("Filename,hasform,isListEntry");
            for fname,listentries,calcfields in zip(target_files,listentries_per_file,calcfields_per_file):
                print(fname,listentries,calcfields);
        if(current_argument=="--remove" or current_argument=="-r"):
            # for now copy the file to a new file and don't replace the original, but at some point we might want to provide a destructive-mode option 
            #print("WARNING: This will modify files in situ. Make a backup before trying this, ok? [type 'YES' to continue]");
            #responsetext = raw_input("OK? ");
            #if(responsetext!="YES"):
            #    print("Operation cancelled.");
            #    sys.exit();
            for target_file in target_files:
                remove_calcfields(target_file);

# vim: ts=4 sw=4 et
