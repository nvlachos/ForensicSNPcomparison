import sys, getopt, glob, os, csv, openpyxl as opxl, xlrd as xr, xlwt as xw
from loci import LocusHelper
from copy import deepcopy
 
# The Sample class is used to contain and control data at the sample level. Each sample object will contain all
# Loci (SNPs & STRs) within the finaldict variable, which is created by formatting the sourcedict according to which source it came from
# The genotype, coverage, major allele frequency and flags of all loci will be stored, as well as sequence and frequency of STR loci
class Sample:
   
   # Constructor for a new sample. Standard initialization requires a name and the instrument/company that the sample was run on 
   def __init__(self, name = "None", sampleType = "None"):
      # The name of the sample. Used for identification
      self.name = name
      # The source of the sample. Options include "Thermo", "MiSeq", and "CLC" at the moment
      self.sampleSource = sampleType
      # A variable that stores what type of loci are being analyzed in thermo samples. It is not used in CLC or MiSeq samples right now
      self.subtype = sampleType
      # A variable that stores which kit was used to create the results contained in the sample
      self.kit = "?"
      # The unfiltered and mostly complete version of the data contained in the source file
      self.sourceDict = {}
      # The edited/filtered universally formated data from the source dictionary 
      self.finalDict = {}
      #These values get assigned in function formatDist from values in locusHelper in loci.py and will be later implemented to
      # help properly flag and make genotype calls for CLC data
      # The minimum coverage for an STR to be called 
      self.minimumSTRCoverage = None
      # The STR interpretation threshold, the value,as a percentage, of total calls necessary to be called as an allele
      self.strIT = None
      # The STR analytical threshold, the value, as a percentage, of total calls necessary to be considered as a possible allele (stutter, contamination, etc.)
      self.strAT = None
      # Minimum coverage for a SNP to be called
      self.minimumSNPCoverage = None
      # The SNP interpretation threshold, the value,as a percentage, of total calls necessary to be called as an allele
      self.snpIT = None
      # The SNP analytical threshold, the value, as a percentage, of total calls necessary to be considered as a possible allele
      self.snpAT = None
      # The value of the major allele frequency, as a percentage,  at which the calling of a heterzygote vs homozygote becomes difficult
      self.hohetCutoff = None
      # The minimum value, as a percentage, that the major allele frequency needs to be to call a homozygote
      self.homomin = None
      # The average coverage of all SNP loci within the sample
      self.averageSNPCoverage = 0
      # The average coverage of all STR loci within the sample
      self.averageSTRCoverage = 0
      ### Variables Currently only used for MiSeq samples, as they are included in the excel report files
      # Name of the project that the sample is from
      self.project = None
      # Analysis version that the sample is from
      self.analysisID = None
      # Name of the run the sample is from
      self.runname = None
      # file creation info of sample from UAS software
      self.created = None
      
   # A function that will add information to flags to help determine why it failed (eg MAF for imbalanced site)
   # flagset is the original values pulled from sample reports
   # values is the full locus entry from the source dictionary
   # Freq is the frequency that the allele appears within in the locus for STRs (or MAF for SNP sites)
   # locusID is the name of current locus being worked on
   # STRtotalCoverage is the coverage of all alleles at that locus for STRs, variabnle is not used for SNPs and
   # is used to discriminate between the two
   def processFlags(self, flagset, values, Freq, locusID, STRtotalCoverage):
      # The string that will be returned as the value for the flag after processing
      returnStringofFlags = ''
      # Checks to see if the incoming flagset is a string or a single element of a list or a list
      # All functions that call this should now be fixed to send only losts, not any strings
      if isinstance(flagset, str):
         # Should no longer Occur
         print("Error; String received as flagset parameter within processFlags")
      # Checks to see if sample came from a Thermo Instrument so that it can properly parse the flags
      if self.sampleSource == "Thermo":
         # All thermop flags begin with an underscore and this removes it
         sepFlags = flagset[0].replace('_', '')
#         print("L:", sepFlags)
         # Looks to see if the allele has more than one flag, which are separated by a semi colon and splits them into a list if more than one
         if ";" in sepFlags:
#            print ("Semi found")
            sepFlags = sepFlags.split(";")[:-1]
         # If there is only one flag, it is turned into a list (rather than a string)
         if isinstance(sepFlags, str):
            sepFlags = [sepFlags]
         # Increments through the list of flags to process each one independantly
         for flag in sepFlags:
            #If the return string already has a length, then a semicolon is added to sepoarate new flag info
            if (len(returnStringofFlags) > 0):
               returnStringofFlags = returnStringofFlags + ";"
            # If the flag is for Major Allele Frequency, then the value is pulled straight from the original report line data rounded to 2 places 
            if flag == "MAF":
               returnStringofFlags = returnStringofFlags + "MAF(" + values[11][0:values[11].find(".")+3] + ")"
            # If the flag is for "NOC"? then the quality score value is pulled from the original report line
            elif flag == "NOC":
               returnStringofFlags = returnStringofFlags + "NOC(" + values[10] + ")"
            # If the flag is for Percent Positive Coverage, then the value is pulled straight from the original report line data rounded to 2 places 
            elif flag == "PPC":
               returnStringofFlags = returnStringofFlags + "PPC(" + values[9][0:values[9].find(".")+3] + ")"
            # If the flag is for Coverage, then the value is pulled straight from the original report line data
            elif flag == "COV":
               returnStringofFlags = returnStringofFlags + "COV(" + values[2] + ")"
            # If the flag is for Above Stochastic Threshold, nothing is done as this is not a flag
            elif flag == "ABOVEST":
               pass
            # If the flag is for Below Peak Height Ratio, then frequency of allele is calculated as well as allele coverage out of total locus coverage
            elif flag == "BELOWPHR":
#               print (values)
               returnStringofFlags = returnStringofFlags + "Below PHR (" + Freq +"%/"+ values[5]+" of " + STRtotalCoverage + ")"
            # If the flag is an empty string, do nothing 
            elif flag == " " or flag == '' or flag == "NA":
               pass
            # If some new or unknown flag is encountered
            else:
               # Unknown Thermo Flag encountered
               print ("New Thermo Flag FOUND:", flag, "!!! in", self.name)
#         print (returnStringofFlags)
      # If the sample source came from MiSeq then process the flags here
      elif self.sampleSource == "MiSeq":
#         print (flagset)
         sepFlags = flagset[0]
#         print("L:", sepFlags)
         # Looks to see if there are multiple flags by searching for a comma delimiter
         if "," in sepFlags:
#            print ("Semi found")
            sepFlags = sepFlags.split(", ")
         # If only one flag exists then it is converted to a list item
         if isinstance(sepFlags, str):
            sepFlags = [sepFlags]
#         print("SF:", sepFlags)
         # Increments through list of flags to process each one independantly
         for flag in sepFlags:
#            print (values)
            #If the return string already has a length, then a semicolon is added to sepoarate new flag info
            if (len(returnStringofFlags) > 0):
               returnStringofFlags = returnStringofFlags + ";"
            # If flag is for interpretation threshold, then STRs get allele frequency percentage and coverage, but a SNP locus gets
            # total coverage and coverage at each base 
            if flag == "interpretation threshold":
               if STRtotalCoverage != None:
                  returnStringofFlags = returnStringofFlags + "IT(" + Freq + "%/" + str(values[2]) + ")"
               else:
                  returnStringofFlags = returnStringofFlags + "IT(" + str(values[2]) + "/{" + str(values[3]) + "," + str(values[4]) + "," + str(values[5]) + "," + str(values[6]) +"})"
            # If flag is for imbalance, then STR gets allele frequnecy and allele coverage of total locus coverage, and SNP gets MAF value
            elif flag == "imbalanced":
               if STRtotalCoverage != None:
                  returnStringofFlags = returnStringofFlags + "MAF(" + Freq + "%/" + str(values[2]) + " of "+ STRtotalCoverage + ")"
               else:
                  returnStringofFlags = returnStringofFlags + "MAF(" + str(Freq) + ")" 
            # If the flag is for allele count, then the allele count is calculated and returned
            elif flag == "allele count":
               returnStringofFlags = returnStringofFlags + "AC(" + str(len(self.getAlleles(locusID))) + ")"
            # If the flag is for stutter, then the allele count is calculated and returned
            elif flag == "stutter":
               returnStringofFlags = returnStringofFlags + "stutter(" + str(len(self.getAlleles(locusID))) + ")"
            # If the flag is an empty string, do nothing 
            elif flag == " " or flag == '':
               pass
            # If a new or unknown flag is encountered
            else:
               # Unknown MiSeq Flag encountered
               print ("New MiSeq Flag FOUND:", flag, "!!!")
      # If the sample is from a CLC report
      elif self.sampleSource == "CLC":
#          print (flagset, values)
          # Increment through the list of flags found for the allele
          for flag in flagset:
             #If the return string already has a length, then a semicolon is added to sepoarate new flag info
             if (len(returnStringofFlags) > 0):
                returnStringofFlags = returnStringofFlags + ";"
            # If the flag is for coverage then the  value from the original report is added  
             if flag == "COV1":
                returnStringofFlags = returnStringofFlags + "COV(" + str(values[0][5]) + ")"
            # If COV2 is found then the second highest allele coverage is added to
             elif flag == "COV2":
                returnStringofFlags = returnStringofFlags + "2nd(" + str(values[0][3]) + " of "+ str(values[0][5]) + ")"
            # If the flag is for MAF then the value is added from the calculated value in SNPMAF
             elif flag == "MAF":
                returnStringofFlags = returnStringofFlags + "MAF(" + str(values[0][4]) + ")"
            # If the flag is an empty string, do nothing 
             elif flag == " " or flag == '':
               pass
            # If a new or unknown flag is encountered
             else:
                # Unknown CLC Flag encountered
               print ("New CLC Flag FOUND:", flag, "!!!")
      # If a new sample cource type is encountered
      else:
         print ("New Sample Source Type\"" + self.sampleSource + "\"seen in process Flags!!!")
#      print (returnStringofFlags)
      return returnStringofFlags  
        
      
      
   # A function that is used to calculate the allele frequencies for SNP loci in CLC samples
   # counts - a 4 item list of the number of base calls [a,c,g,t]
   # returns a list as follows [major allele, major allele coverage, minor allele, minor allele coverage, the major allele frequency, total coverage]
   def SNPMAF(self, counts):
#      print(counts)
      # Creates a lookup list for reference
      bases = ["A","C","G","T"]
      # Sets the initial major to a unique unused character, so it can not be used if a major allele is not found
      first = "Z"
      # Sets the initial major index to a negative number, so it can not be used if a major allele is not found
      firstIndex = -1
      # Sets the inital minor to a unique unused character, so it can not be used if a major allele is not found
      second = "Z"
      # Sets the initial major index to a negative number, so it can not be used if a major allele is not found
      secondIndex = -1
      # Sets the inital value of the major allele frequency to 0
      MAF = 0.00
      # Sets the inital major allele coverage to 0
      firstCoverage = 0
      # Sets the inital minor allele coverage to 0
      secondCoverage = 0
      # For loop that increments through the counts list to find which index has the highest value(coverage)
      for i in range(0,4):
         # Logic to replace the first allele coverage and index, if the new value is higher than the old
         if int(counts[i]) > firstCoverage:
            first = bases[i]
            firstIndex = i
            firstCoverage = int(counts[i])
      # For loop that increments through the counts list to find which index has the second highest value(coverage)
      for i in range(0,4):
         # Logic to replace the second allele coverage and index, if the new value is higher than the old
         if int(counts[i]) > secondCoverage and i != firstIndex:
            second = bases[i]
            secondIndex = i
            secondCoverage = int(counts[i])
      # Calculates the major allele frequency, as a percentage, using the highest coverage value over total coverage
      if ((int(counts[0])+int(counts[1])+int(counts[2])+int(counts[3])) > 0):
         MAF = 100*firstCoverage/(int(counts[0])+int(counts[1])+int(counts[2])+int(counts[3]))
      # returns a list as follows [major allele, major allele coverage, minor allele, minor allele coverage, the major allele frequency, total coverage]
#      print (first, firstCoverage, second, secondCoverage, round(MAF,2), int(counts[0])+int(counts[1])+int(counts[2])+int(counts[3]))
      return [first, firstCoverage, second, secondCoverage, round(MAF,2), int(counts[0])+int(counts[1])+int(counts[2])+int(counts[3])]
   
   # A helper function that adds flags to CLC samples
   # counts - counts - a 4 item list of the number of base calls [a,c,g,t]
   # returns a list of the flags that were generated for the locus
   def clcSNPFlags(self, counts):
      # A list of flags for the sample
      flags = []
      # A MAF flag will be thrown if the major allele frequency falls above the hohetcutoff value but below the minimum homozygote value of the sample
      if (counts[4] > self.hohetCutoff and counts[4] < self.homomin):
         flags.append("MAF")
      # A COV1 flag will be thrown if the coverage of the major allele falls below the minimum SNP coverage set in the sample object
      if (counts[1] > 0 and counts[1] < self.minimumSNPCoverage): 
         flags.append("COV1")
      # A COV2 flag will be thrown if the coverage of the second allele falls below the minimum SNP coverage set in the sample object
      elif (counts[3] > 0 and counts[3] < self.minimumSNPCoverage):
         flags.append("COV2")
      # Returns the full list of flags
      return flags
         
   # A helper function that creates the SNP genotype call from an input line of a CLC file
   # clcline is the line read in from a clc report file
   # returns a list of [list of major allele, major coverage, minor allele, minor coverage, MAF, total coverage], genotype, [flags]]
   def clccalc(self, clcline):
      # Determines if the current locus is a Y SNP (therefore only having one allele) by looking up the SNP in the LocusHelper object
#      print (clcline[0])
      if (clcline[0] in LocusHelper.ySNPs):
#            print (clcline[0], "found in Y list")
            isY = True
      else:
            isY = False
      # Creates the major/minor allele summary list
      counts = self.SNPMAF(clcline[4:8])
      # Creates the list of flags for the locus
      flags = self.clcSNPFlags(counts)
      # Sets the inital value of the genotype as blank, in case there is an unhandled error
      genotype = None
#      print (clcline)
      # If the sample has not been flagged
      if (len(flags) == 0):
         # If the locus is on the Y chromosome
         if isY:
            # The largest genotype becomes the major allele
            genotype = counts[0]
         # If the locus is on any chromosome, other than Y
         else:
            #---------- Add theshold calculations in here
            # If there is a second allele above 0 coverage
             if(counts[4] > self.homomin):
               genotype = counts[0]+counts[0]
             else:
               genotype = counts[0]+counts[2]
            
            
            
#            if(counts[3] > 0):
#               # Genotype becomes a heterozygote call
#               genotype = counts[0]+counts[2]
#            else:
#               # Genotype becomes a homozygous call
#               genotype = counts[0]+counts[0]
      # If the sample has one flag
      elif (len(flags) == 1):
         # If the locus is on a Y-Chromosome
         if isY:
            # If the single flag is for MAF
            if flags[0] == "MAF":
               # Genotype call is the major allele
               genotype = counts[0]
               # Flag is changed to indicate that since it is a Y SNP and there are more than 1 allele, this might be another contributor or contamination
               flags[0] = "MAF (multi-male?)"
            # If the single flag is for coverage 1
            elif flags[0] == "COV1":
               # The genotype is changed to N
               genotype = "N"
            # If the single flag is for coverage 2
            elif flags[0] == "COV2":
               #---------- The genotype is still kept as the major allele, no other action is taken for low second allele coverage (should it?)
               genotype = counts[0]
            # Catch all in case flags are created later that are forgotten to put in this list
            else:
               print("Unknown Y-MAF flag found in flags array in clccalc")
         # If the locus is from any chromosome, other than Y
         else:
            # If the single flag is MAF
            if flags[0] == "MAF":
               # The genotype becomes a heterozygote
               # More extensive selection will need to be done here if used formally
#               print ("This!", counts[0]+counts[2])
               if(counts[4] > self.homomin):
                  genotype = counts[0]+counts[0]
               else:
                  genotype = counts[0]+counts[2]
            # If the single flag is for coverage 1
            elif flags[0] == "COV1":
               # The genotype becomes NN
               genotype = "NN"
            # If the single flag is for coverage 2
            elif flags[0] == "COV2":
               # ---------- The genotype is recorded as homozygote, more action may be required though
               genotype = counts[0]+counts[0]
            # Catch all in case flags arer created later that are forgotten to put in this list
            else:
               print("Unknown A-MAF flag found in flags array in clccalc")
      # If there are 2 flags for a sample (only 2 possibilities right now, MAF&COV1 and MAF&COV2)
      elif (len(flags) == 2):
         # If the locus is on a Y chromosome
         if isY:
            # If the second flag is for coverage 1
            if flags[1] == "COV1":
               # The genotype becomes N
               genotype = "N"
               # The MAF flag is changed to indicate that there may be contamination or a second contributor
               flags[0] = "MAF (multi-male?)"
            # If the second flag is for coverage 2
            elif flags[1] == "COV2":
               #---------- Genotype becomes the major allele, should do something for 2nd coverage?
               genotype = counts[0]
            # Catch all for any other flags that may be created later that are not placed in the list here
            else:
               print("Unknown Y-COV flag found in flags array in clccalc")
         # If the locus is on any chromosome, other than Y
         else:
            # If the second flag is for coverage 1
            if flags[1] == "COV1":
               # Genotype is changed to NN
               genotype = "NN"
            # If the second flag is for coverage 2
            elif flags[1] == "COV2":
               #---------- The genotype is recorded as a homozygote, more action may be required though
               if(counts[4] > self.homomin):
                  genotype = counts[0]+counts[0]
               else:
                  genotype = counts[0]+counts[2]
            # Catch all for any other flags that may be created later that are not placed in the list here
            else:
               print("Unknown A-COV flag found ("+flags[1]+") in flags array in clccalc")
#      print (counts, genotype, flags)
      # To ensure correct sizing, an if statement checking the size of the flags variable is used
      # The allele information (counts), and the genotype call are returned along with the list of flags for the locus (an empty string is
      # returned if there are no flags
      if (len(flags) == 0):
          return [counts, genotype, ['']]
      else:
          return [counts, genotype, flags]
         

     
   # This helper function transforms all source dictionaries into a standard format to allow universal access in downstream processing 
   def formatDict(self):
#      print (self.name)
      # A list of active loci names found in the sample
      locs = []
      # Checks the sample source to see what instrument/software produced it. Each source will require different formatting.
      # If the source came as output from CLC
      if self.sampleSource == "CLC":
         # Increment through all loci within the source dictionary
         for keys,values in self.sourceDict.items():
#            print (values[10])
            # Checks if the locus has any coverage at all
            if int(values[10])>0:
               # Adds the locus to the active list
               locs.append(keys)
            else:
               #print (keys, "did not have any reads")
               # Continues to the next locus if this locus did not have any reads
               pass
#            print (locs, len(locs))
         # Calls the source function to identify what kit the sample was run through
         self.kit = source(locs)
#         print (self.kit)
         # If the original run was from a Thermo product then the appropriate coverage and thresholds are applied to the sample
         if self.kit[0] == "Thermo":
            self.minimumSTRCoverage = LocusHelper.thermominimumSTRCoverage
            self.strIT = LocusHelper.thermostrIT
            self.strAT = LocusHelper.thermostrAT
            self.minimumSNPCoverage = LocusHelper.thermominimumSNPCoverage
            self.snpIT = LocusHelper.thermosnpIT
            self.snpAT = LocusHelper.thermosnpAT
            self.hohetCutoff = LocusHelper.thermohohetCutoff
            self.homomin = LocusHelper.thermohomommin
         # If the original run was from a MiSeq product then the appropriate coverage and thresholds are applied to the sample
         elif self.kit[0] == "MiSeq":
            self.minimumSTRCoverage = LocusHelper.miseqminimumSTRCoverage
            self.strIT = LocusHelper.miseqstrIT
            self.strAT = LocusHelper.miseqstrAT
            self.minimumSNPCoverage = LocusHelper.miseqminimumSNPCoverage
            self.snpIT = LocusHelper.miseqsnpIT
            self.snpAT = LocusHelper.miseqsnpAT
            self.hohetCutoff = LocusHelper.miseqhohetCutoff
            self.homomin = LocusHelper.miseqhomomin
         # If the original run was not from a thermo or MiSeq product, an unknown CLC source error is printed
         else:
            print("Unknown CLC source in formatDict from sample: " + self.name)
            return
         # The final formatting is done entry by entry in the source dictionary by updating the values of the key
         # by using the appropriate elements. The type (SNP) is added followed by the genotype, coverage, MAF, and flags
         for keys,values in self.sourceDict.items():
#            print (values[10])
            # If the allele have the necessary minimum coverage to be called
            if int(values[10])>=self.minimumSNPCoverage:
                  # Processes the list of values using the clccalc function to format original numbers to the universal format
#                  print(keys)
                  allele = self.clccalc(values)
#                  print ("A",allele)
                  # Processes the flags for the allele
                  flag = self.processFlags(allele[2], allele, allele[0][4], keys, None)
                  # Updates the dictionary entry for the allele in the final dict
                  self.finalDict.update({keys:["SNP"]+[allele[1]]+[allele[0][5]]+[allele[0][4]]+[flag]})
            # Allele did not have any reads      
            else:
#               print (keys, "did not have any reads")
               pass
          
      #---------- If the sample source is Thermo (also applies to S5, it just is not included    
      elif self.sampleSource == "Thermo":
         # Increments through all keys and values in the source dictionary
         for keys,values in self.sourceDict.items():
#            print(" Source to format:", keys,values)
            # Each locus is added to the active SNP list, since no extraneous loci will be included in the report
            locs.append(keys)
            # If the loci indicate that the panel was for SNPs (by using the identifylocustype heklper function)
            if identifyLocusType(keys) == "SNP":
               # Flags are processed to add more detailed information
               flag = self.processFlags([values[12]], values, values[11], keys,  None)
               # Values of keys are updated using the first, second, eleventh, and twelfth values from the source dictionary for the key,
               # which correlate to the genotype, coverage, MAF, and flags
               self.finalDict.update({keys:["SNP"]+[values[1]]+[values[2]]+[values[11]]+[flag]})
            # If the loci indicate that the panel was for STRs (by using the identifylocustype heklper function)
            elif identifyLocusType(keys) == "STR":
               # values for the keys are updated using the sorted output from the strFreq helper function (which standardizes output)
               self.finalDict.update({keys:["STR",sorted(self.strFreq(keys,values))]})
            else: #---------- Locus was not identified in list; therefore type is unknown and no further processing will occur
               pass
         # Uses source function to identify the original 
         self.kit = source(locs)
         if self.kit[0] == "Unknown Manufacturer":
            print ("Unknown Kit source based on percentages")
            self.toString()
      # If the sample source came from a MiSeq
      elif self.sampleSource == "MiSeq":
         # Increment through all keys and values in the source dictionary
         for keys,values in self.sourceDict.items():
            # Add all loci to the active loci list
            locs.append(keys)
            # If the current key is a SNP
            if keys in LocusHelper.allSNPs:
               # Creates an empty variable to hold return values from SNPMAF function
               valuesfromSNPMAF = []
#               print("Sending to SNPMAF:", values)
               # Sends the base coverages to SNPMAF for processing
               valuesfromSNPMAF = self.SNPMAF(values[3:7])
#               print("last source dict check before format:", keys, "|", values)
#               print("Received SNPMAF:", valuesfromSNPMAF)
               # Processes the flags to provide more details
               flag = self.processFlags(values[7], values, valuesfromSNPMAF[4], keys, None)
               # Updates the entry for the current key with the processed genotype, coverage, MAF, and flag info
               self.finalDict.update({keys:["SNP"]+[values[1]]+[values[2]]+[valuesfromSNPMAF[4]]+[flag]})
#               print ("dic check:", self.finalDict.get(keys))
            # If the current key is an STR
            elif keys in LocusHelper.allSTRs:
#               print ("Checking here:", values[1])
               # Creates a temp variable that is made up of the return values from the STRFreq function
               tempList = ["STR"]+[self.strFreq(keys, values[1])]
#               print ("Updated: ", keys, tempList)
               # Updates the key with temp variable
               self.finalDict.update({keys:tempList})
            # Determine which kit the sample was run through using the source function
         # Calls source after finished to determine which kit was used to create the run
         self.kit = source(locs)
         if self.kit[0] == "Unknown Manufacturer":
            print ("Unknown Kit source based on percentages")
            self.toString()
      # If an unknown source is encountered
      else:
         print("Unknown sample source type in formatDict")
#      Prints all keys/values for a sample
#      for keys,values in self.finalDict.items():
#          print(keys, values)
      self.getCounts()
   
   # Function to determine average SNP and STR coverage for the sample   
   def getCounts(self):
      # Initiation of variables
      totalSNPs = 0
      totalSNPcount = 0
      totalSTRs = 0
      totalSTRcount = 0
      # Goes through each locus in the in the sample
      for locus, values in self.finalDict.items():
#         print ("349:", locus, values)
         # If the locus is a SNP then just add the coverage from the value
         if (values[0] == "SNP"):
            # Increment the total number of SNPs found for proper averaging
            totalSNPs = totalSNPs + 1
            totalSNPcount = totalSNPcount + int(values[2])
         # If the locus is an STR, then each allele has to be summed individually
         elif(values[0] == "STR"):
            # Increment the total number of STRs found for proper averaging
            totalSTRs = totalSTRs + 1
            for allele in values[1]:
#               print("356:", allele, allele[1], type(allele[1]), int(float(allele[1])), type(int(float(allele[1]))), totalSTRcount)
               totalSTRcount = totalSTRcount + int(float(allele[1]))
#               print (totalSTRcount)
         # If an unknown locus type (not SNP or STR) is encountered while summing and averaging
         else:
            print ("Unknown locus type in getCounts()")
#      print("360:", totalSNPcount, totalSNPs, totalSTRcount, totalSTRs)
      # Averages the total SNP count b ythe total nuymber of SNPs
      if (totalSNPs != 0):
         self.averageSNPCoverage = round(totalSNPcount/totalSNPs)
      # Skips if the total number of SNPs is 0
      else:
         pass#print("No SNP loci in", self.name)
      # Averages the total STR count b ythe total nuymber of STRs
      if (totalSTRs != 0):
         self.averageSTRCoverage = round(totalSTRcount/totalSTRs)
      # Skips if the total number of STRs is 0   
      else:
         pass #print("No STR loci in", self.name)
#      print("369:",self.name, self.averageSNPCoverage, self.averageSTRCoverage)

   # Used to check if string can be an int or not (for checking allele id (i.e. 12(int) vs. 12.0(float)))
   # Became a necessary due to naming of alleles in different platforms and importing, sometimes its 9 and sometimes 9.0, so now
   # we need to check occaisonally
   def RepresentsInt(self, s):
       try: 
           int(s)
           return True
       except ValueError:
           return False
         
   # Is a function to find the most common repeat words for a sequence, however this will be a much larger
   # undertaking and will be returned to at a later time
   def sequenceCounter(self, sequence, repeatSize, alleleSize):
#      print (sequence)
#     # Only checks for word sizes of 3,4, or5
      if repeatSize < 3 or repeatSize > 5:
         return
      # If word size is in correct range then continue
      else:
         # Creates an empty word list to hold all unique 
         words = []
         # Goes through and checks next k-mer word to see if has been added to list yet
         for i in  range(0, len(sequence)-repeatSize):
            if sequence[i:i+repeatSize] not in words:
               words.append(sequence[i:i+repeatSize])
#         print (words)
         # Creates a sister list that holds the corresponding counts of the words in the first list
         wordCounts = []
         # Increments through words list to count how many times it appears in the sequence
         for i in words:
            wordCounts.append(sequence.count(i))
#         print (wordCounts)
         # Finds the highest value in the count list
         maxCount = max(wordCounts)
         # Finds all k-mers that match the max size value
         for i in range(0, len(wordCounts)):
            if (wordCounts[i] == maxCount):
#               print("Biggest " + str(repeatSize) + ":", words[i])
               pass
      
   # A helper function for Thermo STR files that calculates total coverage and frequencies for a single locus
   # listofAlleles - input that is a list in source format of alleles for a locus
   # Returns the same list of alleles, but in the standard format [allele, coverage, sequence, MAF, Frequency of this allele, flags]
   def strFreq(self, locusID, listOfAlleles):
   #   print (listOfAlleles)
      # The total coverage for this locus
      totalCoverage = 0
      # The coverage of the most prevalent allele
      largestAlleleCoverage = 0
      # List of improved and standardized alleles for this locus
      improvedListOfAlleles = []
      # Goes through all alleles for a locus
      for allele in listOfAlleles:
#         print("STRFREQ:", allele)
         # Adds to the total coverage for the locus
         if (self.sampleSource == "Thermo"):
            totalCoverage = totalCoverage + int(allele[5])
            # Checks to see if this allle is the new largest allele
            if (int(allele[5]) > largestAlleleCoverage):
               largestAlleleCoverage = int(allele[5])
         elif (self.sampleSource == "MiSeq"):
#            print ("!", allele)
            totalCoverage = totalCoverage + int(allele[2])
            # Checks to see if this allle is the new largest allele
            if (int(allele[2]) > largestAlleleCoverage):
               largestAlleleCoverage = int(allele[2])
      
      # Calculates the MAF for the locus
      if (totalCoverage > 0):
         maf = round(100*largestAlleleCoverage/totalCoverage, 2)
      else:
         maf = 0
#      print("Total coverage:", totalCoverage, " | ", largestAlleleCoverage, " | ", maf)
      # Goes through all alleles in the locus, again, to check for flags and to standardize the alleles
      for allele in listOfAlleles:
#         print("0",allele[0], "1", allele[1], "2", allele[2], "3", allele[3], "4", allele[4], "5", allele[5], "6", allele[6])
         if self.sampleSource == "Thermo":
            # Creates a string based on using processed flags 
            flag  = self.processFlags([allele[4]], allele, str(round(100*int(allele[5])/totalCoverage, 2)), locusID, str(totalCoverage))
            # Creates a list of info for this particular allele that is as follows [allele, coverage, sequence, MAF, Frequency of this allele, flags]
            tempEntry = [allele[3], allele[5], allele[6], maf, round(100*int(allele[5])/totalCoverage, 2), flag]
#            print ("TE:", tempEntry)
            # Adds the tempEntry into the improved alleles list
            improvedListOfAlleles.append(tempEntry)
         elif self.sampleSource == "MiSeq":
            # Creates a string based on using processed flags 
            flag = self.processFlags(allele[5], allele,str(round(100*int(allele[2])/totalCoverage, 2)), locusID, str(totalCoverage))
            #  Creates a temp variable to add to the improved alleles list. The logic is to chcek if it is a standard allele or a plus
            # If it can be an int then the decimal and 0 are removed, if it is a float the other characters are kept
            if self.RepresentsInt(allele[0]):
               tempEntry = [str(allele[0])[:-2], str(allele[2]), allele[3], maf, round(100*int(allele[2])/totalCoverage, 2), flag]
            else:
               tempEntry = [str(allele[0]), str(allele[2]), allele[3], maf, round(100*int(allele[2])/totalCoverage, 2), flag]   
#            print ("TempEntry:", tempEntry)
            # Adds the tempEntry into the improved alleles list
            improvedListOfAlleles.append(tempEntry)
            
#      print ("ILOA from STRFreq:", improvedListOfAlleles)
      # Returns the improved list of alleles
      return improvedListOfAlleles
   
   
   
   # Calculates the average coverage for a sample
   # Returns a list of [average coverage, average SNP coverage, average STR coverage]
   def getAverage(self):
      snpsumandn = [0,0]
      strsumandn = [0,0]
      # Performs summing and count of each locus of a sample
      for keys,values in self.finalDict.items():
         # If the locus is a SNP, it just adds the current coverage to total SNP sum and incrments n
          if values[0] == "SNP":
            snpsumandn[0] = snpsumandn[0] + values[2]
            snpsumandn[1] = snpsumandn[1] + 1
         # If the locus is a STR, it goes through each allele and adds the voverage to the total sum and increments n       
          elif values[0] == "STR":
            for allele in values[1:]:
               strsumandn[0] = strsumandn[0] + allele[1]
               strsumandn[1] = strsumandn[1] + 1
      # Calculates the averages for each type (in percentages)
      averageSNPcoverage = 100*snpsumandn[0]/snpsumandn[1]
      averageSTRcoverage = 100*strsumandn[0]/snpsumandn[1]
      averageTOTALcoverage = 100*(snpsumandn[0]+strsumandn[0])/(snpsumandn[1]+strsumandn[1])
      # Returns the list of each type of average, rounded to 2 decimal places
      return [round(averageTOTALcoverage, 2), round(averageSNPcoverage,2), round(averageSTRcoverage,2)]
   
   
   # Returns the allele info for a given locus 
   def getAlleles(self, locusID):
      return self.finalDict.get(locusID, [0, "Error:",locusID, "Not Found"])[1]
   
   # Returns all similar alleles for a locus (e.g. all 12.2's)
   def getSpecificAllele(self, locusID, alleleID):
      # Creates a temp variable to hold all found alleles
      tempAlleles = []
#      print("Looking up", alleleID, type(alleleID), "for", locusID, type(locusID))
      # Increments through all alleles for a locus
      for alleles in self.finalDict.get(locusID, "No"+locusID+"found in sample"+self.name)[1]:
#         print ("Checking if same:", alleles[0], type(alleles[0]), alleleID, type(alleleID))
         # Add to the list if it matches
         if alleles[0] == alleleID:
             tempAlleles.append(alleles)
#      print("Size is:", len(tempAlleles))
      # If there are any alleles found that match the reqauested ID then they are returned
      if len(tempAlleles)>0:
         return tempAlleles
      # Returns None because no alleles were found
      else:
         return None
      
   # Helper function to print out all variables contained within a sample
   def toString(self):
      # Prints the general variables contained within all samples
      print("Sample name:",self.name)
      print("Sample source:",self.sampleSource)
      print("Sample sub-type:",self.subtype)
      print("Sample kit:", self.kit)
      print("Minimum STR coverage:", self.minimumSTRCoverage)
      print("STR Interpretation Threshold:", self.strIT)
      print("STR Analytical Threshold:", self.strAT)
      print("Minimum SNP coverage:", self.minimumSNPCoverage)
      print("SNP Interpretation Threshold:", self.snpIT)
      print("SNP Analytical Threshold:", self.snpAT)
      print("Clear Heterozygote Cutoff:", self.hohetCutoff)
      print("Clear Homozygote Cutoff:", self.homomin)
      print("Average SNP coverage:", self.averageSNPCoverage)
      print("Average STR Coverage:", self.averageSTRCoverage)
      print("Number of Source loci (not total alleles):", len(self.sourceDict))
      #---------- Prints the contents of the source dictionary for debugging purposes
      if (False):
         for locus,values in self.sourceDict.items():
               print("Source:", locus, values)
               if values[0] == "STR":
                  for locus2 in values[1]:
                     print ("Allele:", locus, locus2)
               elif values[0] == "SNP":
                  print("Genotype:", values[1:])
      print ("Number of Final loci (not total alleles):", len(self.finalDict))
      #---------- Prints the contents of the final dictionary for debugging purposes
      if (False):
         for locus,values in self.finalDict.items():
            print("Final:", locus, values)
            if values[0] == "STR":
               for locus2 in values[1]:
                  print ("Allele:", locus, locus2)
            elif values[0] == "SNP":
               print("Genotype:", values[1:])
     # Prints a line to show the end of the particular sample
      print ("End sample:", self.name)
      
#________________________________________END SAMPLE CLASS______________________________________________________#




# The main method of the genocount script
def main(argv):
   # Creates an empty default input folder variable
   # example - C:\Users\filth\OneDrive\Documents\CFSRU\16413\test
   inputfolder = ''
   # Creates an empty default output folder variable
   # example - C:\Users\filth\OneDrive\Documents\CFSRU\16413\test
   outputfile = ''
   # Creates a list of booleans, defaulted to false, that track which info the user is interested in
   showCoverage = False
   showGenotypes = False
   showFlags = False
   showMAF = False
   showSNPs = False
   showSTRs = False
   # A list that contains a list of all sample objects created for analysis
   samples = []
   # Parses the command line to sort the arguments
   try:
      opts, args = getopt.getopt(argv,"hcgftmnsi:o:",["ifile=","ofile="])
   except getopt.GetoptError:
      sys.exit(2)
   # Checks which options were requested from the command line
   for opt, arg in opts:
      # The h option gives the general syntax of how to order the command line arguments, and then exits
      if opt == '-h':
         print ('help - genocount.py -i <inputfile> -o <outputfile> -c -g -f -m -n -t')
         sys.exit()
      # If -i or --ifile, the argument contains where to find the folder that contains all the files to be analyzed
      elif opt in ("-i", "--ifile"):
         inputfolder = os.path.normpath(arg)
      # If -o or --ofile, the argument contains where to put the output files that were analyzed   
      elif opt in ("-o", "--ofile"):
         outputfile = arg
      # Sets the boolean to show coverages of samples in the output
      elif opt == '-c':
         showCoverage = True
      # Sets the boolean to show genotypes of the samples in the output
      elif opt == '-g':
         showGenotypes = True
      # Sets the boolean to show flags of the samples in the output   
      elif opt == '-f':
         showFlags = True
      # Sets the boolean to show the MAFs of the samples in the output
      elif opt == '-m':
         showMAF = True
      # Sets the boolean to show SNPs of the samples in the output
      elif opt == '-n':
         showSNPs = True
      # Sets the boolean to show STRs of the samples in the output
      elif opt == '-t':
         showSTRs = True
   # Testing arguments      
#   print ('Input folder is:',inputfolder)
#   print ('Output file is:',outputfile)
#   print ("Show Coverage? ", showCoverage)
#   print ("Show Genotypes? ", showGenotypes)
#   print ("Show Flags? ", showFlags)
#   print ("Show MAF? ", showMAF)
#   print ("Show SNPs? ", showSNPs)
#   print ("Show STRs? ", showSTRs)

#   if outputfile == '':
#      outputfile = os.path.join
   
   # Goes through each file that is located within the input folder designated as a command line argument
   for filename in glob.glob(inputfolder + "/*.*"):
      # Creates a dummy sample object
      temp = Sample()
#      print (filename)
      # If the filename has an .xlsx extension (thereby being from the MiSeq), begin MiSeq analysis
      if (os.path.splitext(filename)[1] == '.xlsx'):
#           print("Miseq placeholder1")
#          temp = miseqParser(filename)
           samples.append(parseMiSeqXLSX(filename))
      elif (os.path.splitext(filename)[1] == '.xls'):
#           print("Miseq placeholder2")
#          temp = miseqParser(filename)
           samples.append(parseMiSeqXLS(filename))
      # If the filename has a .csv extension (thereby being from a Thermo instrument), begin Thermo analysis
      elif (os.path.splitext(filename)[1]) == '.csv':
          # Creates a temporary variable to hold the return value of pgmParser, which could be a list or a sample object
          # depending on which kit it was run through
          temp = pgmParser(filename)
          # If temp is a sample then pgmParser determined it to be a single sample of SNPs
          if isinstance(temp, Sample):
            #print("Found a SNP Thermo")
            # Adds the temp sample to the samples list
            samples.append(pgmParser(filename))
         # If temp is a list, then the pgmParser determined the samples to be STRs
          elif isinstance(temp, list):
#           print("Size of samples =", len(temp))
#           for i in temp:
#              print(i.name)
            # Increments through all samples in the list (created from the STR file) and adds each one to the sample list
            for i in temp:
                samples.append(i)
          # pgmParser can not identify what the input type is
          else:
            print("pgmParser is returning unknown type:", type(temp))
       # If the filename has a .tsv extension (thereby being a CLC file), begin CLC analysis     
      elif (os.path.splitext(filename)[1]) == '.tsv':
#         temp = clcParser(filename)
         # Appends the output from the clcParser function (which will be a sample object) to the samples list
          samples.append(clcParser(filename))
      # If an extension other than csv,tsv, or .xlsx is encountered
      else:
         # The filename and then extension is printed for debugging
         print(os.path.splitext(filename)[0])
         print(os.path.splitext(filename)[1] + ": Extension Not Found")
   # Increments through all samples in the list to see what has been added (Debugging use only)
#   for i in samples:
#       i.toString()
   # Sends the sample list, which is sorted based on sample name, to the export function
   csvexport(sorted(samples, key=lambda x: x.name, reverse=True), showCoverage,showGenotypes,showFlags,showMAF,showSNPs,showSTRs, os.path.join(inputfolder, "output", outputfile))
   
 # Helper function that is passed a Thermo file to be parsed and converted into samples
 # fileToParse - the filename of the thermo output file to be converted
 # returns a single sample object for SNP panels or a list of samples for STR Panels from Thermo instruments
def pgmParser(fileToParse):
#   print (fileToParse)
   # Creates an empty variable to hold a sample object, if the sample was run through a SNP panel
   newSample = None
   # File object that opens the Thermo file to be able to read
   pgm_file_object = open(fileToParse, 'r')
   # Creates a CSV reader object from the open file 
   csv_pgm = csv.reader(pgm_file_object)
   # Row counter for csv reader
   count = 0
   # Variable to hold all lines of an STR Thermo sample, due to the structure of the file this list is necessary to match samples together
   STRLines = None
   # Boolean to show if the file is for SNPs or STRs
   isSNP = True
   # Reads each line from the csv reader (of the input file)
   for row in csv_pgm:
      # Checks the length of the first line
      if (count == 0):
 #        print(row)
         # If the length of the first line is 1 or 7, then the file is STRs
         if (len(row) == 1 or len(row) == 7):
            # Grabs the 'run name' from the first line of the file and adds it as the first element of the STRLines list
            STRLines = [row[0][row[0].index("=")+1:]]
            # Sets the isSNP to false to tell the rest of the processing that this sample is full of STRs
            isSNP = False
         # If the length of the first line is 16, then the sample is of SNPs
         elif (len(row) == 16):
            # Creates a new sample object with the sample filename and 'Thermo' as constructor arguments
            newSample = Sample(os.path.basename(fileToParse)[:-4], "Thermo")
            # Sets the sample subtype to 'SNP'
            newSample.subtype = "SNP"
         # If the length of the first line is not 1,7, or 16 then it is unknown format
         else:
            #---------- Prints that an unknown Thermo CSV has been encountered, do more
            print("Unknown CSV file")
      # After the first line check, if the sample is SNPs
      elif (isSNP):
         # Update the source dictionary value of the locus to the remaining portion of the line (includes counts, + cov, - cov, ppc, GQ, flags)
         newSample.sourceDict.update({row[2]: [row[2]]+row[4:]})
      # After the first line check, if the sample is STRs
      else:
         # Skips the first 6 rows of the STR file (it is meta data)
         if (count > 6):
            # Adds the row (allele) to the STRLines list
            STRLines.append(row)
      # Increments the line counter variable   
      count = count + 1
#        Prints all keys and values in the samples source dictionary
#      for keys,values in newSample.sourceDict.items():
#          print(keys)
#          print(values)
   # If the sample has been tagged as SNPs
   if (isSNP):
      # Call the format dictionary function
      newSample.formatDict()
      # Returns the complete and formatted sample 
      return newSample
   # If the sample is STRs
   else:
      # Returns the converted STRLines variable which was parsed into seperate samples using the pgmSTRparser function
      return pgmSTRparser(STRLines)      

# Converts a list of STR alleles from multiple samples into seperate sample objects
# linesOfSamples - a list of STR allleles that was read in from an STR run on a Thermo machine
# Returns a list of finalized samples that have been parsed and seperated into sample objects from the original conglomeration of samples/alleles
def pgmSTRparser(linesOfSamples):
   # Pulls out the run name from the linesOfSamples variable
   runName = linesOfSamples[0]
   # Creates the default empty variable types that will be used to track and organize the different sampels within the list
   # Keeps the name of the current sample available to know when a new one is encountered (source list is sorted by sample name)
   currentSampleName = ""
   # Creates a temp current sample object to assign to when the end of a sample is reached
   currentSample = None
   # Keeps track of the current locus to know when a new one is encountered (source list is sorted by locus, after sample name)
   currentLocusName = ""
   # Keeps track of the allleles for the current loci for the current sample
   currentLocus = []
   # Contains all the samples that have been found and sorted 
   samples = []
   # Removes the run name from the list
   del linesOfSamples[0]
   # Begins going through the list line by line
   for alleles in linesOfSamples:
#      print("XXX", alleles)
      #If sample name is not the same as last allele processed, 
      if (str(alleles[1]) != currentSampleName):
         # Closes out the previous sample if there is one
         if(currentSample != None):
#            print("New sample! Closing out...", currentSampleName)
#            currentSample.toString()
            # Tells the currentSample to format its dictionary
            currentSample.formatDict()
            # Adds the current sample to the list of samples
#            currentSample.toString()
            samples.append(currentSample)
#         print("Creating new sample...", str(alleles[1]))
         # Creates a new sample using data from line
         currentSample = Sample(str(alleles[1]), "Thermo")
         currentSampleName = str(alleles[1])
         currentSample.subtype = "STR"
         currentLocusName =""
         currentLocus = []
      #If the locus is NOT the same as the last one processed
      if (str(alleles[2]) != currentLocusName):
         # If this is not the first locus in a sample then the previous locus (and all alleles) is updated in the source dict of the sample object
         if (currentLocusName != ""):
#            print("new locus in", currentSampleName, "closing out...", currentLocusName)
#            print("Old Locus name:", currentLocusName, ";New Locus name:", str(alleles[2]))
#            print("newEntry:", currentLocusName, newEntry)
#            print("dic before", len(currentSample.sourceDict))
            # Sets the source dictionary of the current locus to the value stored as current locus which is a list of all alleles for that locus
            currentSample.sourceDict[currentLocusName] = currentLocus
#            print(currentSample.sourceDict.get(currentLocusName, "Not Found"))
#            print("dic after", len(currentSample.sourceDict))
#         print("Creating new locus of ", str(alleles[2]), "in", currentSampleName)
         # Changes current locus name to the new one from the line
         currentLocusName = str(alleles[2])
         # Empties out the alleles list to refresh for the new locus
         currentLocus = []
#      print("Adding", currentLocusName, "to", currentSampleName)
      # Adds the current locus 
      currentLocus.append(alleles)
#      print ("Alleles:", alleles)
#   print ("NewEntry:",newEntry)
   # Sets the source dictionary of the current locus to the value stored as current locus which is a list of all alleles for that locus (Ending call)
   currentSample.sourceDict[currentLocusName] = currentLocus
   # Tells the currentSample to format its dictionary (Ending call)
   currentSample.formatDict()
   # Adds the current sample to the list of samples
   samples.append(currentSample)
#   samples[0].toString()
   # Returns the list of finalized samples 
   return samples
   
# Helper function to determine more information from a sample that was sent through the CLC pipeline
# fileToParse - the whole file that will be parsed into a sample
# Returns a finalized sample object 
def clcParser(fileToParse):
#    print (fileToParse)
   # Creates a default sample object using the filename and CLC as arguments
    newSample = Sample(os.path.basename(fileToParse)[:-4], "CLC")
    # Creates a file reader to open the filename that was passed in
    pgm_file_object = open(fileToParse, 'r')
    # Creates a CSV reader from the file reader
    tsv_clc = csv.reader(pgm_file_object, delimiter='\t')
    # Creates a row counter
    count = 0
    # ---------- ???
    numOfTens = 0
    # Go through every row of the file
    for row in tsv_clc:
        # Sets the name to blank
        name = ""
        # If the locus name contains a dash (For the oddball 10mer of the Phenotypic MiSeq SNPs)
        # Example N29insA-rs1805005-rs1805006-rs2228479-rs11547464-rs1805007-rs201326893_Y152OCH-rs1110400-rs1805008-rs885479_M_P_chr16_89985663_()_89986241_Hg19
        if row[0].find("-") >= 0:
#            print (row[0], "has dash!", numOfTens)
            # Gets the position of all dashes found in the locus name
            dashIndices = [pos for pos, char in enumerate(row[0]) if char == '-']
            # Goes through all instances of dashes encountered in the source locus name
            for i in range(0, len(dashIndices)):
               # Offset the index of the dash to the next letter
               dashIndices[i] = dashIndices[i] + 1
#            print (row[0].find("_", row[0].find("_") + 1))

            dashIndices = [0] + dashIndices + [row[0].find("_", row[0].find("_") + 1) + 1]
            # Sequentially grabs the locus name of the locus between each set of dashes 
            name = row[0][dashIndices[numOfTens]:dashIndices[numOfTens+1]-1]
#            print(dashIndices[numOfTens], dashIndices[numOfTens+1]-1, name)
            numOfTens = numOfTens + 1
         # If the locus name does not contain a dash, but it contains an open bracket within the first 5 characters (For the very short Thermo SNPs)
        elif row[0].find("(") <= 5:
            # The locus name is from the start of the original locus name to index of the open bracket
           name = row[0][0:row[0].find("(")]
         # The locus name does not have a dash, or an open bracket within the first 5 chars (All other regular SNPs)
        else:
            # The locus name becomes everything up to the first underscore
           name = row[0][0:row[0].find("_")]
        # Update the value of the locus name with the data in the row from the csv reader
#        print (name)
        newSample.sourceDict.update({name: [name] + row[1:]})
        # Increment the row counter
        count = count + 1
#    for keys,values in newSample.sourceDict.items():
#        print(keys)
#        print(values)
   # Formats the new sample dictionary
    newSample.formatDict()
   # Returns the sample
    return newSample

#  Helper function to parse out samples from MiSeq machine
# fileToParse - file to parse into a sample
# Return a finalized sample
def miseqParser(fileToParse):
#    print(fileToParse)
#    print ("Test1: ", fileToParse[-4:], "Test2: ", fileToParse[-5:])
   # If the file has an ".XLS" extension it is sent to the XLS parser
    if fileToParse[-4:] == ".xls":
#      print("Processing XLS File:", fileToParse)
      return parseMiSeqXLS(fileToParse)
   # If the file has an ".XLSX" extension it is sent to the XLSX parser
    elif fileToParse[-5:] == ".xlsx":
#      print("Processing XLSX File:", fileToParse)
      return parseMiSeqXLSX(fileToParse)
    # Unknown extension encountered
    else:
      print("Unknown file extension encountered in miseqParser", fileToParse)
    
#  Helper function to parse out samples from MiSeq machine for XLS file
# fileToParse - file to parse into a sample (will be .XLS)
# Return a finalized sample     
def parseMiSeqXLS(fileToParse):
#   print (fileToParse)
   # Uses xlrd to open the workbook file
   workbook_in = xr.open_workbook(fileToParse, formatting_info=True, on_demand=True)
#   print (workbook_in.sheet_names()) 
   sample = Sample("Default", "MiSeq")
   #All the cell locations to start and stop iterations (values will be decremented by one later due to 0/1 start positions)
   ySTRUpperStart = 15-1
   ySTRUpperStop = 38-1
   xSTRUpperStart = 15-1
   xSTRUpperStop = 21-1
   aSTRUpperStart = 15-1
   aSTRUpperStop = 42-1
   iSNPUpperStart = 15-1
   iSNPUpperStop = 108-1
   apSNPUpperStart = 14-1
   apSNPUpperStop = 67-1
   ySTRLowerStart = 43-1
   xSTRLowerStart = 26-1
   aSTRLowerStart = 47-1
   iSNPLowerStart = 113-1
   apSNPLowerStart = 71-1
   # The Stop positions will depend on the number of alleles found and will be limited by the nrows function of worksheet from xlrd
   
   # Checks to see if there is a sheet named "SNP Data" which would indicate a Ancestry and Phenotypic report
   if "SNP Data" in workbook_in.sheet_names():
#      print ("Found A&P File")
      # Sets the active sheet to the "SNP Data" worksheet
      worksheet = workbook_in.sheet_by_name("SNP Data")
      # Goes through each row of the sheet and processes it accordingly
      for rowNumber in range(0,worksheet.nrows):
         # If the first cell contains "Sample" then the sample name is set to the second cell
         if worksheet.cell(rowNumber,0).value == "Sample":
            sample.name = os.path.split(fileToParse)[1]+"("+(worksheet.cell(rowNumber,1).value)+")"
#            print (worksheet.cell(rowNumber,1).value)
        # If the first cell contains "Project" then the sample project name is set to the second cell
         if worksheet.cell(rowNumber,0).value == "Project":
            sample.project = worksheet.cell(rowNumber,1).value
#            print (worksheet.cell(rowNumber,1).value)
         # If the first cell contains "Analysis" then the sample analysis name is set to the second cell
         if worksheet.cell(rowNumber,0).value == "Analysis":
            sample.analysisID = worksheet.cell(rowNumber,1).value
#            print (worksheet.cell(rowNumber,1).value)
         # If the first cell contains "Run" then the sample run name is set to the second cell
         if worksheet.cell(rowNumber,0).value == "Run":
            sample.runname = worksheet.cell(rowNumber,1).value
#            print (worksheet.cell(rowNumber,1).value)
         # If the first cell contains "Created" then the sample created variable is set to the second cell
         if worksheet.cell(rowNumber,0).value == "Created":
            sample.created = worksheet.cell(rowNumber,1).value
#            print (worksheet.cell(rowNumber,1).value)
         # Initiates source dictionaries from top half of miseq excel sheet. Worksheet is laid out in 3 sets of columns, hence the triple
         # copy of tyhe coding
         if rowNumber >= apSNPUpperStart and rowNumber <= apSNPUpperStop:
             if (worksheet.cell(rowNumber, 0).value != xr.empty_cell.value) and (worksheet.cell(rowNumber, 1).value != "INC"):
                sample.sourceDict[worksheet.cell(rowNumber, 0).value] = [worksheet.cell(rowNumber, 0).value, str(worksheet.cell(rowNumber, 1).value[:1])+str(worksheet.cell(rowNumber, 1).value[-1:]),0,0,0,0,0,[worksheet.cell(rowNumber, 2).value]]
#                print("H/E:",sample.sourceDict.get(worksheet.cell(rowNumber, 0).value))
               
             if (worksheet.cell(rowNumber, 4).value != xr.empty_cell.value) and (worksheet.cell(rowNumber, 5).value != "INC"):
                sample.sourceDict[worksheet.cell(rowNumber, 4).value] = [worksheet.cell(rowNumber, 4).value, str(worksheet.cell(rowNumber, 5).value[:1])+str(worksheet.cell(rowNumber, 5).value[-1:]),0,0,0,0,0,[worksheet.cell(rowNumber, 6).value]]
#                 print("Common:",sample.sourceDict.get(worksheet.cell(rowNumber, 4).value))
               
             if (worksheet.cell(rowNumber, 8).value != xr.empty_cell.value) and (worksheet.cell(rowNumber, 9).value != "INC"):
                sample.sourceDict[worksheet.cell(rowNumber, 8).value] = [worksheet.cell(rowNumber, 8).value, str(worksheet.cell(rowNumber, 9).value[:1])+str(worksheet.cell(rowNumber, 9).value[-1:]),0,0,0,0,0,[worksheet.cell(rowNumber, 10).value]]
#                print("BGA:",sample.sourceDict.get(worksheet.cell(rowNumber, 8).value))
                
         # Modifies source dictionary reads
         if rowNumber >= apSNPLowerStart:
             # Checks if the allele has any reads (does not update anything if there are no reads)
             if (worksheet.cell(rowNumber, 3).value != xr.empty_cell.value and worksheet.cell(rowNumber, 3).value != 0):
               # creates a temp variable to hold current value of the allele from source dict
                tempDictValue = sample.sourceDict.get(worksheet.cell(rowNumber, 0).value)
#                print("TDV:",sample.sourceDict.get(worksheet.cell(rowNumber, 0).value))
                # Checks for each base coverage and updates as necessary
#                print ("!!!", worksheet.cell(rowNumber, 3).value)
#                print (tempDictValue)
                if tempDictValue != None:
                  if (worksheet.cell(rowNumber, 1).value == "A" or worksheet.cell(rowNumber, 1).value == "a"):
                     tempDictValue[3] = int(worksheet.cell(rowNumber, 3).value)
  #                   sample.sourceDict[worksheet.cell(rowNumber, 0).value] = tempDictValue
                  if (worksheet.cell(rowNumber, 1).value == "C" or worksheet.cell(rowNumber, 1).value == "c"):
                     tempDictValue[4] = int(worksheet.cell(rowNumber, 3).value)
  #                   sample.sourceDict[worksheet.cell(rowNumber, 0).value] = tempDictValue
                  if (worksheet.cell(rowNumber, 1).value == "G" or worksheet.cell(rowNumber, 1).value == "g"):
                     tempDictValue[5] = int(worksheet.cell(rowNumber, 3).value)
  #                   sample.sourceDict[worksheet.cell(rowNumber, 0).value] = tempDictValue
                  if (worksheet.cell(rowNumber, 1).value == "T" or worksheet.cell(rowNumber, 1).value == "t"):
                     tempDictValue[6] = int(worksheet.cell(rowNumber, 3).value)
                  # Sums all base coverages to create a totalcoverage
                  tempDictValue[2] = tempDictValue[3]+tempDictValue[4]+tempDictValue[5]+tempDictValue[6]
                  # Sets the source dict to the updated value
                  sample.sourceDict[worksheet.cell(rowNumber, 0).value] = tempDictValue
  #                print("PDV:", sample.sourceDict.get(worksheet.cell(rowNumber, 0).value))
   
   # Checks to see if an "iSNPs" sheet exists in the workbook, indicating an STR and identity report       
   elif "iSNPs" in workbook_in.sheet_names():
#      print ("Found STR and iSNP")
      # Sets the current worksheet to iSNPs
      currentWorksheet = workbook_in.sheet_by_name("iSNPs")
      # Goes through each row in the sheet and processes them appropriately
      for rowNumber in range(0,currentWorksheet.nrows):
#         print (rowNumber, "of", currentWorksheet.nrows)
         # If the first cell contains "Sample" then the sample name is set to the second cell
         if currentWorksheet.cell(rowNumber,0).value == "Sample":
            sample.name = os.path.split(fileToParse)[1]+"("+(currentWorksheet.cell(rowNumber,1).value)+")"
#            print (currentWorksheet.cell(rowNumber,1).value)
        # If the first cell contains "Project" then the sample project name is set to the second cell
         if currentWorksheet.cell(rowNumber,0).value == "Project":
            sample.project = currentWorksheet.cell(rowNumber,1).value
#            print (currentWorksheet.cell(rowNumber,1).value)
         # If the first cell contains "Analysis" then the sample analysis name is set to the second cell
         if currentWorksheet.cell(rowNumber,0).value == "Analysis":
            sample.analysisID = currentWorksheet.cell(rowNumber,1).value
#            print (currentWorksheet.cell(rowNumber,1).value)
         # If the first cell contains "Run" then the sample run name is set to the second cell
         if currentWorksheet.cell(rowNumber,0).value == "Run":
            sample.runname = currentWorksheet.cell(rowNumber,1).value
#            print (currentWorksheet.cell(rowNumber,1).value)
         # If the first cell contains "Created" then the sample created variable is set to the second cell
         if currentWorksheet.cell(rowNumber,0).value == "Created":
            sample.created = currentWorksheet.cell(rowNumber,1).value
#            print (worksheet.cell(rowNumber,1).value)
         # Initiates source dictionaries from top half of miseq excel sheet. 
         if rowNumber >= iSNPUpperStart and rowNumber <= iSNPUpperStop:
            # If the cell is not empty then add new locus to sample source dictionary
            if (currentWorksheet.cell(rowNumber, 0).value != xr.empty_cell.value):
               sample.sourceDict[currentWorksheet.cell(rowNumber, 0).value] = [currentWorksheet.cell(rowNumber, 0).value, str(currentWorksheet.cell(rowNumber, 1).value[:1])+str(currentWorksheet.cell(rowNumber, 1).value[-1:]),0,0,0,0,0,[currentWorksheet.cell(rowNumber, 2).value]]
#               print("iSNP:",sample.sourceDict.get(currentWorksheet.cell(rowNumber, 0).value))
         # Updates all the coverage values from loci found in the upper half
         if rowNumber >= iSNPLowerStart:
            # Checks if the allele has any reads (does not update anything if there are no reads)
            if (currentWorksheet.cell(rowNumber, 3).value != xr.empty_cell.value and currentWorksheet.cell(rowNumber, 3).value != 0):
               tempDictValue = sample.sourceDict.get(currentWorksheet.cell(rowNumber, 0).value)
#               print("TDV:",sample.sourceDict.get(currentWorksheet.cell(rowNumber, 0).value))
               if (currentWorksheet.cell(rowNumber, 1).value == "A" or currentWorksheet.cell(rowNumber, 1).value == "a"):
                  tempDictValue[3] = currentWorksheet.cell(rowNumber, 3).value
#                  sample.sourceDict[currentWorksheet.cell(rowNumber, 0).value] = tempDictValue
               if (currentWorksheet.cell(rowNumber, 1).value == "C" or currentWorksheet.cell(rowNumber, 1).value == "c"):
                  tempDictValue[4] = currentWorksheet.cell(rowNumber, 3).value
#                  sample.sourceDict[currentWorksheet.cell(rowNumber, 0).value] = tempDictValue
               if (currentWorksheet.cell(rowNumber, 1).value == "G" or currentWorksheet.cell(rowNumber, 1).value == "g"):
                  tempDictValue[5] = currentWorksheet.cell(rowNumber, 3).value
#                  sample.sourceDict[currentWorksheet.cell(rowNumber, 0).value] = tempDictValue
               if (currentWorksheet.cell(rowNumber, 1).value == "T" or currentWorksheet.cell(rowNumber, 1).value == "t"):
                  tempDictValue[6] = currentWorksheet.cell(rowNumber, 3).value   
               tempDictValue[2] = tempDictValue[3]+tempDictValue[4]+tempDictValue[5]+tempDictValue[6]
               sample.sourceDict[currentWorksheet.cell(rowNumber, 0).value] = tempDictValue
#               print("PDV:", sample.sourceDict.get(currentWorksheet.cell(rowNumber, 0).value))
      
      # Swicthes to Y-STR sheet to extract all STRs on this sheet     
      currentWorksheet = workbook_in.sheet_by_name("Y STRs")
      # Creates an empty dictionary to hold all discovered STRs
      Alleles = {}
      # Increment through all rows in the sheet
      for rowNumber in range(0,currentWorksheet.nrows):
        #If the row number falls within the upper section boundaries then create a new locus in dictionary and use flags and calls
        if rowNumber >= ySTRUpperStart and rowNumber <= ySTRUpperStop:
            # If the value of the cell is empty or "INC", then do nothing with the locus, else add new locus and info to dictionary
            if (currentWorksheet.cell(rowNumber, 0).value != xr.empty_cell.value and currentWorksheet.cell(rowNumber, 2).value != "INC"):
               Alleles[currentWorksheet.cell(rowNumber, 0).value] = [currentWorksheet.cell(rowNumber, 1).value, currentWorksheet.cell(rowNumber, 2).value]
#               print("Temp Dict entry made for A STR:", currentWorksheet.cell(rowNumber, 0).value,"as", Alleles.get(currentWorksheet.cell(rowNumber, 0).value))
        # If the rownumber falls within the lower section boundaries, update locus in dictionary with allele, coverage, typed?, and sequence info 
        if rowNumber >= ySTRLowerStart:
            # Checks if the allele has any reads (does not update anything if there are no reads)
            if (currentWorksheet.cell(rowNumber, 3).value != xr.empty_cell.value and currentWorksheet.cell(rowNumber, 3).value != 0):
#               print ("Sample check:", currentWorksheet.cell(rowNumber, 0).value, sample.sourceDict.get(currentWorksheet.cell(rowNumber, 0).value))
               # Adds locus to dictionary
               if (sample.sourceDict.get(currentWorksheet.cell(rowNumber, 0).value) == None):
#                   print ("No locus found yet in sample sourceDict. Checking Allele to begin locus with:", currentWorksheet.cell(rowNumber, 0).value, Alleles.get(currentWorksheet.cell(rowNumber, 0).value))
                   sample.sourceDict[currentWorksheet.cell(rowNumber, 0).value] = ["STR", [[currentWorksheet.cell(rowNumber, 1).value,currentWorksheet.cell(rowNumber, 2).value,currentWorksheet.cell(rowNumber, 3).value,currentWorksheet.cell(rowNumber, 4).value, Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[0], [Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[1]]]]]
               # Updates the entry for the locus in the dictionary
               else:
                   oldLocusArray = sample.sourceDict.get(currentWorksheet.cell(rowNumber, 0).value)
#                   print ("Trying to append:", [currentWorksheet.cell(rowNumber, 1).value,currentWorksheet.cell(rowNumber, 2).value,currentWorksheet.cell(rowNumber, 3).value,currentWorksheet.cell(rowNumber, 4).value, Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[0], Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[1]], "to", oldLocusArray)
#                   print ("Test1:", oldLocusArray[0], oldLocusArray[1])
                   oldLocusArray[1].append([currentWorksheet.cell(rowNumber, 1).value,currentWorksheet.cell(rowNumber, 2).value,currentWorksheet.cell(rowNumber, 3).value,currentWorksheet.cell(rowNumber, 4).value, Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[0], [Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[1]]])
#                   print ("Test2:", oldLocusArray[1])
#                   print ("New entry:", newLocusArray)
                   sample.sourceDict[currentWorksheet.cell(rowNumber, 0).value] = oldLocusArray
      
      # Swicthes to Y-STR sheet to extract all STRs on this sheet         
      currentWorksheet = workbook_in.sheet_by_name("X STRs")
      # Creates an empty dictionary to hold all discovered STRs
      Alleles = {}
      # Increment through all rows in the sheet
      for rowNumber in range(0,currentWorksheet.nrows):
         #If the row number falls within the upper section boundaries then create a new locus in dictionary and use flags and calls
         if rowNumber >= xSTRUpperStart and rowNumber <= xSTRUpperStop:
            # If the value of the cell is empty or "INC", then do nothing with the locus, else add new locus and info to dictionary
            if (currentWorksheet.cell(rowNumber, 0).value != xr.empty_cell.value and currentWorksheet.cell(rowNumber, 2).value != "INC"):
               Alleles[currentWorksheet.cell(rowNumber, 0).value] = [currentWorksheet.cell(rowNumber, 1).value, currentWorksheet.cell(rowNumber, 2).value]
#               print("Temp Dict entry made for A STR:", currentWorksheet.cell(rowNumber, 0).value,"as", Alleles.get(currentWorksheet.cell(rowNumber, 0).value))
         # If the rownumber falls within the lower section boundaries, update locus in dictionary with allele, coverage, typed?, and sequence info 
         if rowNumber >= xSTRLowerStart:
            # Checks if the allele has any reads (does not update anything if there are no reads)
            if (currentWorksheet.cell(rowNumber, 3).value != xr.empty_cell.value and currentWorksheet.cell(rowNumber, 3).value != 0):
#               print ("Sample check:", currentWorksheet.cell(rowNumber, 0).value, sample.sourceDict.get(currentWorksheet.cell(rowNumber, 0).value))
               # Adds locus to dictionary
               if (sample.sourceDict.get(currentWorksheet.cell(rowNumber, 0).value) == None):
#                   print ("No locus found yet in sample sourceDict. Checking Allele to begin locus with:", currentWorksheet.cell(rowNumber, 0).value, Alleles.get(currentWorksheet.cell(rowNumber, 0).value))
                   sample.sourceDict[currentWorksheet.cell(rowNumber, 0).value] = ["STR", [[currentWorksheet.cell(rowNumber, 1).value,currentWorksheet.cell(rowNumber, 2).value,currentWorksheet.cell(rowNumber, 3).value,currentWorksheet.cell(rowNumber, 4).value, Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[0], [Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[1]]]]]
               # Updates the entry for the locus in the dictionary
               else:
                   oldLocusArray = sample.sourceDict.get(currentWorksheet.cell(rowNumber, 0).value)
#                   print ("Trying to append:", [currentWorksheet.cell(rowNumber, 1).value,currentWorksheet.cell(rowNumber, 2).value,currentWorksheet.cell(rowNumber, 3).value,currentWorksheet.cell(rowNumber, 4).value, Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[0], Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[1]], "to", oldLocusArray)
#                   print ("Test1:", oldLocusArray[0], oldLocusArray[1])
                   oldLocusArray[1].append([currentWorksheet.cell(rowNumber, 1).value,currentWorksheet.cell(rowNumber, 2).value,currentWorksheet.cell(rowNumber, 3).value,currentWorksheet.cell(rowNumber, 4).value, Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[0], [Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[1]]])
#                   print ("Test2:", oldLocusArray[1])
#                   print ("New entry:", newLocusArray)
                   sample.sourceDict[currentWorksheet.cell(rowNumber, 0).value] = oldLocusArray
                   
      # Swicthes to Y-STR sheet to extract all STRs on this sheet            
      currentWorksheet = workbook_in.sheet_by_name("Autosomal STRs")
      # Creates an empty dictionary to hold all discovered STRs
      Alleles = {}
      # Increment through all rows in the sheet
      for rowNumber in range(0,currentWorksheet.nrows):
        # If the row number falls within the upper section boundaries then create a new locus in dictionary and use flags and calls
        if rowNumber >= aSTRUpperStart and rowNumber <= aSTRUpperStop:
            # If the value of the cell is empty or "INC", then do nothing with the locus, else add new locus and info to dictionary
            if (currentWorksheet.cell(rowNumber, 0).value != xr.empty_cell.value and currentWorksheet.cell(rowNumber, 2).value != "INC"):
               Alleles[currentWorksheet.cell(rowNumber, 0).value] = [currentWorksheet.cell(rowNumber, 1).value, currentWorksheet.cell(rowNumber, 2).value]
#               print("Temp Dict entry made for A STR:", currentWorksheet.cell(rowNumber, 0).value,"as", Alleles.get(currentWorksheet.cell(rowNumber, 0).value))
        # If the rownumber falls within the lower section boundaries, update locus in dictionary with allele, coverage, typed?, and sequence info 
        if rowNumber >= aSTRLowerStart:
            # Checks if the allele has any reads (does not update anything if there are no reads)
            if (currentWorksheet.cell(rowNumber, 3).value != xr.empty_cell.value and currentWorksheet.cell(rowNumber, 3).value != 0):
#               print ("Sample check:", currentWorksheet.cell(rowNumber, 0).value, sample.sourceDict.get(currentWorksheet.cell(rowNumber, 0).value))
               # Adds locus to dictionary
               if (sample.sourceDict.get(currentWorksheet.cell(rowNumber, 0).value) == None):
#                   print ("No locus found yet in sample sourceDict. Checking Allele to begin locus with:", currentWorksheet.cell(rowNumber, 0).value, Alleles.get(currentWorksheet.cell(rowNumber, 0).value))
                   sample.sourceDict[currentWorksheet.cell(rowNumber, 0).value] = ["STR", [[currentWorksheet.cell(rowNumber, 1).value,currentWorksheet.cell(rowNumber, 2).value,currentWorksheet.cell(rowNumber, 3).value,currentWorksheet.cell(rowNumber, 4).value, Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[0], [Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[1]]]]]
               # Updates the entry for the locus in the dictionary
               else:
                   oldLocusArray = sample.sourceDict.get(currentWorksheet.cell(rowNumber, 0).value)
#                   print ("Trying to append:", [currentWorksheet.cell(rowNumber, 1).value,currentWorksheet.cell(rowNumber, 2).value,currentWorksheet.cell(rowNumber, 3).value,currentWorksheet.cell(rowNumber, 4).value, Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[0], Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[1]], "to", oldLocusArray)
#                   print ("Test1:", oldLocusArray[0], oldLocusArray[1])
                   oldLocusArray[1].append([currentWorksheet.cell(rowNumber, 1).value,currentWorksheet.cell(rowNumber, 2).value,currentWorksheet.cell(rowNumber, 3).value,currentWorksheet.cell(rowNumber, 4).value, Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[0], [Alleles.get(currentWorksheet.cell(rowNumber, 0).value)[1]]])
#                   print ("Test2:", oldLocusArray[1])
#                   print ("New entry:", newLocusArray)
                   sample.sourceDict[currentWorksheet.cell(rowNumber, 0).value] = oldLocusArray
                   
   # Creates a list of locus ids only to use in determining the kit used
   tempLoci = []
   # Goes through all source samples to get locus names
   for key, value in sample.sourceDict.items():
#      print (" XLS Locus:", key, "Value", value)
      tempLoci.append(key)
   # Sets the kit variable using the loci present in the list
#   sample.kit = source(tempLoci)[1]
   # Debugging sample print statement
   if (False):
         sample.toString()
   sample.formatDict()
   if (False):
         sample.toString()
   # Returns the sample to the caller
   return sample
   
      
def parseMiSeqXLSX(fileToParse):
   # Uses openXL to open the workbook file
   workbook_in = opxl.load_workbook(fileToParse)
   sample = Sample("Default", "MiSeq")
   #All the cell locations to start and stop iterations (values not decremented by one in openxlpy)
   ySTRUpperStart = 15
   ySTRUpperStop = 38
   xSTRUpperStart = 15
   xSTRUpperStop = 21
   aSTRUpperStart = 15
   aSTRUpperStop = 42
   iSNPUpperStart = 15
   iSNPUpperStop = 108
   apSNPUpperStart = 14
   apSNPUpperStop = 67
   ySTRLowerStart = 43
   xSTRLowerStart = 26
   aSTRLowerStart = 47
   iSNPLowerStart = 113
   apSNPLowerStart = 71
   # The Stop positions will depend on the number of alleles found and will be limited by the nrows function of worksheet from xlrd
   
   # Checks to see if the workbook contains a "SNP Data" page indicating a Ancestry and Phenotype report file
   if "SNP Data" in workbook_in.get_sheet_names():
#      print ("Found A&P File")
      # Sets active sheet to "SNP Data" sheet
      worksheet = workbook_in["SNP Data"]
#      print (len(worksheet['A']))
      # Goes through every row on the SNP Data page and processes line appropriately
      for rowNumber in range(1,(len(worksheet['A'])+1)):
#         print (worksheet.cell(row=rowNumber, column=1).value)
         # If the first cell contains "Sample" then the sample name is set to the second cell
         if worksheet.cell(row=rowNumber, column=1).value == "Sample":
            sample.name = os.path.split(fileToParse)[1]+"("+(worksheet.cell(row = rowNumber,column=2).value)+")"
#            print (worksheet.cell(row=rowNumber, column=1).value)
         # If the first cell contains "Project" then the sample name is set to the second cell
         if worksheet.cell(row=rowNumber, column=1).value == "Project":
            sample.project = (worksheet.cell(row = rowNumber,column=2).value)
#            print (worksheet.cell(row=rowNumber, column=1).value)
         # If the first cell contains "Analysis" then the sample name is set to the second cell
         if worksheet.cell(row=rowNumber, column=1).value == "Analysis":
            sample.analysisID = (worksheet.cell(row = rowNumber,column=2).value)
#            print (worksheet.cell(row=rowNumber, column=1).value)
         # If the first cell contains "Run" then the sample name is set to the second cell
         if worksheet.cell(row=rowNumber, column=1).value == "Run":
            sample.runname = (worksheet.cell(row = rowNumber,column=2).value)
#            print (worksheet.cell(row=rowNumber, column=1).value)
         # If the first cell contains "Created" then the sample name is set to the second cell
         if worksheet.cell(row=rowNumber, column=1).value == "Created":
            sample.created = (worksheet.cell(row = rowNumber,column=2).value)
#            print (worksheet.cell(row=rowNumber, column=1).value)
         # Initiates source dictionaries from top half of miseq excel sheet. Worksheet is laid out in 3 sets of columns, hence the triple
         # copy of tyhe coding
         if rowNumber >= apSNPUpperStart and rowNumber <= apSNPUpperStop:
#             print(rowNumber)
             if (worksheet.cell(row = rowNumber, column=1).value != None):
                sample.sourceDict[worksheet.cell(row = rowNumber, column=1).value] = [worksheet.cell(row = rowNumber, column=1).value, str(worksheet.cell(row = rowNumber, column=2).value)[:1]+str(worksheet.cell(row = rowNumber, column=2).value)[-1:],0,0,0,0,0,[worksheet.cell(row=rowNumber, column=3).value]]
#                print("H/E:",sample.sourceDict.get(worksheet.cell(row = rowNumber, column=1).value))
             if (worksheet.cell(row = rowNumber, column=5).value != None):
                sample.sourceDict[worksheet.cell(row = rowNumber, column=5).value] = [worksheet.cell(row = rowNumber, column=5).value, str(worksheet.cell(row = rowNumber, column=6).value[:1])+str(worksheet.cell(row = rowNumber, column=6).value[-1:]),0,0,0,0,0,[worksheet.cell(row=rowNumber, column=7).value]]
#                print("Common:",sample.sourceDict.get(worksheet.cell(row = rowNumber, column=5).value))
             if (worksheet.cell(row = rowNumber, column=9).value != None):
                sample.sourceDict[worksheet.cell(row = rowNumber, column=9).value] = [worksheet.cell(row = rowNumber, column=9).value, str(worksheet.cell(row = rowNumber, column=10).value[:1])+str(worksheet.cell(row = rowNumber, column=10).value[-1:]),0,0,0,0,0,[worksheet.cell(row=rowNumber, column=11).value]]
#                print("BGA:",sample.sourceDict.get(worksheet.cell(row = rowNumber, column=9).value))
         # Modifies source dictionary reads
         if rowNumber >= apSNPLowerStart:
#             print(rowNumber)
             # Checks if the allele has any reads (does not update anything if there are no reads)
             if (worksheet.cell(row = rowNumber, column=4).value != xr.empty_cell.value and worksheet.cell(row = rowNumber, column=4).value != 0):
                tempDictValue = sample.sourceDict.get(worksheet.cell(row = rowNumber, column=1).value)
#                print("TDV:",sample.sourceDict.get(worksheet.cell(row = rowNumber, column = 1).value))
                # Updates all the coverage values from loci found in the upper half
                if (worksheet.cell(row = rowNumber, column=2).value == "A" or worksheet.cell(row = rowNumber, column=2).value == "a"):
                   tempDictValue[3] = worksheet.cell(row = rowNumber, column=4).value
                   sample.sourceDict[worksheet.cell(row = rowNumber, column=1).value] = tempDictValue
                if (worksheet.cell(row = rowNumber, column=2).value == "C" or worksheet.cell(row = rowNumber, column=2).value == "c"):
                   tempDictValue[4] = worksheet.cell(row = rowNumber, column=4).value
                   sample.sourceDict[worksheet.cell(row = rowNumber, column=1).value] = tempDictValue
                if (worksheet.cell(row = rowNumber, column=2).value == "G" or worksheet.cell(row = rowNumber, column=2).value == "g"):
                   tempDictValue[5] = worksheet.cell(row = rowNumber, column=4).value
                   sample.sourceDict[worksheet.cell(row = rowNumber, column=1).value] = tempDictValue
                if (worksheet.cell(row = rowNumber, column=2).value == "T" or worksheet.cell(row = rowNumber, column=2).value == "t"):
                   tempDictValue[6] = worksheet.cell(row = rowNumber, column=4).value
                   sample.sourceDict[worksheet.cell(row = rowNumber, column=1).value] = tempDictValue
                # Totals all coverages to get a total coverage count
                tempDictValue[2] = tempDictValue[3]+tempDictValue[4]+tempDictValue[5]+tempDictValue[6]
                # Updates the locus in the source dictionary with the coverages
                sample.sourceDict[worksheet.cell(row = rowNumber, column = 1).value] = tempDictValue
#                print("PDV:", sample.sourceDict.get(worksheet.cell(row = rowNumber, column = 1).value))
   
   # Checks to see if an "iSNPs" sheet exists in the workbook, indicating an STR and identity report              
   elif "iSNPs" in workbook_in.get_sheet_names():
#      print ("Found STR and iSNP")
      # Sets the active sheet to "iSNPs"
      currentWorksheet = workbook_in["iSNPs"]
      # Incrememnts through all rows in the sheet
      for rowNumber in range(1,(len(currentWorksheet['A'])+1)):
         # If the first cell contains "Sample" then the sample name is set to the second cell
         if worksheet.cell(row=rowNumber, column=1).value == "Sample":
            sample.name = os.path.split(fileToParse)[1]+"("+(worksheet.cell(row = rowNumber,column=2).value)+")"
#            print (worksheet.cell(row=rowNumber, column=1).value)
         # If the first cell contains "Project" then the sample name is set to the second cell
         if worksheet.cell(row=rowNumber, column=1).value == "Project":
            sample.project = (worksheet.cell(row = rowNumber,column=2).value)
#            print (worksheet.cell(row=rowNumber, column=1).value)
         # If the first cell contains "Analysis" then the sample name is set to the second cell
         if worksheet.cell(row=rowNumber, column=1).value == "Analysis":
            sample.analysisID = (worksheet.cell(row = rowNumber,column=2).value)
#            print (worksheet.cell(row=rowNumber, column=1).value)
         # If the first cell contains "Run" then the sample name is set to the second cell
         if worksheet.cell(row=rowNumber, column=1).value == "Run":
            sample.runname = (worksheet.cell(row = rowNumber,column=2).value)
#            print (worksheet.cell(row=rowNumber, column=1).value)
         # If the first cell contains "Created" then the sample name is set to the second cell
         if worksheet.cell(row=rowNumber, column=1).value == "Created":
            sample.created = (worksheet.cell(row = rowNumber,column=2).value)
#            print (currentWorksheet.cell(row = rowNumber, column=2).value)
         # Initiates source dictionaries from top half of miseq excel sheet
         if rowNumber >= iSNPUpperStart and rowNumber <= iSNPUpperStop:
            if (currentWorksheet.cell(row = rowNumber, column=1).value != None):
               sample.sourceDict[currentWorksheet.cell(row = rowNumber, column=1).value] = [currentWorksheet.cell(row = rowNumber, column=1).value, str(currentWorksheet.cell(row = rowNumber, column=2).value[:1])+str(currentWorksheet.cell(row = rowNumber, column=2).value[-1:]),0,0,0,0,0,[currentWorksheet.cell(row = rowNumber, column=3).value]]
#               print("iSNP:",sample.sourceDict.get(currentWorksheet.cell(row = rowNumber, column=1).value))
         # Modifies source dictionaries with info from the lower half of the excel sheet
         if rowNumber >= iSNPLowerStart:
            # Checks if the allele has any reads (does not update anything if there are no reads)
            if (currentWorksheet.cell(row = rowNumber, column=4).value != xr.empty_cell.value or currentWorksheet.cell(row = rowNumber, column=4).value != 0):
               tempDictValue = sample.sourceDict.get(currentWorksheet.cell(row = rowNumber, column=1).value)
               # Updates all the base coverage numbers
               if (currentWorksheet.cell(row = rowNumber, column=2).value == "A" or currentWorksheet.cell(row = rowNumber, column=2).value == "a"):
                  tempDictValue[3] = currentWorksheet.cell(row = rowNumber, column=4).value
                  sample.sourceDict[currentWorksheet.cell(row = rowNumber, column=1).value] = tempDictValue
               if (currentWorksheet.cell(row = rowNumber, column=2).value == "C" or currentWorksheet.cell(row = rowNumber, column=2).value == "c"):
                  tempDictValue[4] = currentWorksheet.cell(row = rowNumber, column=4).value
                  sample.sourceDict[currentWorksheet.cell(row = rowNumber, column=1).value] = tempDictValue
               if (currentWorksheet.cell(row = rowNumber, column=2).value == "G" or currentWorksheet.cell(row = rowNumber, column=2).value == "g"):
                  tempDictValue[5] = currentWorksheet.cell(row = rowNumber, column=4).value
                  sample.sourceDict[currentWorksheet.cell(row = rowNumber, column=1).value] = tempDictValue
               if (currentWorksheet.cell(row = rowNumber, column=2).value == "T" or currentWorksheet.cell(row = rowNumber, column=2).value == "t"):
                  tempDictValue[6] = currentWorksheet.cell(row = rowNumber, column=4).value
                  sample.sourceDict[currentWorksheet.cell(row = rowNumber, column=1).value] = tempDictValue
               # Sums base coverage numbers to create a total coverage amount
               tempDictValue[2] = tempDictValue[3]+tempDictValue[4]+tempDictValue[5]+tempDictValue[6]
               # Updates source dictionary entry with newest values
               sample.sourceDict[currentWorksheet.cell(row = rowNumber, column = 1).value] = tempDictValue
#               print("PDV:", sample.sourceDict.get(currentWorksheet.cell(row = rowNumber, column = 1).value))
      
      # Switches current sheet to "Y STRs"     
      currentWorksheet = workbook_in["Y STRs"]
      # Creates an empty dictionary to hold all discovered STRs
      Alleles = {}
      # Increment through all rows in the sheet
      for rowNumber in range(1,len(currentWorksheet['A'])):
#        print (rowNumber)
         # If the rownumber falls within the upper section bounds
        if rowNumber >= ySTRUpperStart and rowNumber <= ySTRUpperStop:
            # If the first cell is not empty create a new dictionary entry
            if (currentWorksheet.cell(row = rowNumber, column=1).value != None and currentWorksheet.cell(row = rowNumber, column=3).value != "INC"):
               Alleles[currentWorksheet.cell(row = rowNumber, column=1).value] = [currentWorksheet.cell(row = rowNumber, column=2).value, currentWorksheet.cell(row = rowNumber, column=3).value]
#               print("Temp Dict entry made for Y STR:", currentWorksheet.cell(row = rowNumber, column=1).value,"as", Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value))
        # If the rownumber falls within the lower sections bounds, then update the corresponding locus values in the dictionary
        if rowNumber >= ySTRLowerStart:
            # Checks if the allele has any reads (does not update anything if there are no reads)
            if (currentWorksheet.cell(row = rowNumber, column=4).value != xr.empty_cell.value and currentWorksheet.cell(row = rowNumber, column=4).value != 0):
#               print ("Sample check:", currentWorksheet.cell(row = rowNumber, column=1).value, sample.sourceDict.get(currentWorksheet.cell(row = rowNumber, column=1).value))
               # Adds locus to dictionary
               if (sample.sourceDict.get(currentWorksheet.cell(row = rowNumber, column=1).value) == None):
#                   print ("No locus found yet in sample sourceDict. Checking Allele to begin locus with:", currentWorksheet.cell(row = rowNumber, column=1).value, Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value))
                   sample.sourceDict[currentWorksheet.cell(row = rowNumber, column=1).value] = ["STR", [[currentWorksheet.cell(row = rowNumber, column=2).value,currentWorksheet.cell(row = rowNumber, column=3).value,currentWorksheet.cell(row = rowNumber, column=4).value,currentWorksheet.cell(row = rowNumber, column=5).value, Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[0], [Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[1]]]]]
               # Updates the entry for the locus in the dictionary
               else:
                   oldLocusArray = sample.sourceDict.get(currentWorksheet.cell(row = rowNumber, column=1).value)
#                   print ("Trying to append:", [currentWorksheet.cell(row = rowNumber, column=2).value,currentWorksheet.cell(row = rowNumber, column=3).value,currentWorksheet.cell(row = rowNumber, column=4).value,currentWorksheet.cell(row = rowNumber, column=5).value, Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[0], Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[1]], "to", oldLocusArray)
#                   print ("Test1:", oldLocusArray[0], oldLocusArray[1])
                   oldLocusArray[1].append([currentWorksheet.cell(row = rowNumber, column=2).value,currentWorksheet.cell(row = rowNumber, column=3).value,currentWorksheet.cell(row = rowNumber, column=4).value,currentWorksheet.cell(row = rowNumber, column=5).value, Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[0], [Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[1]]])
#                   print ("Test2:", oldLocusArray[1])
#                   print ("New entry:", newLocusArray)
                   sample.sourceDict[currentWorksheet.cell(row = rowNumber, column=1).value] = oldLocusArray
               
      # Switch to the "X STR" sheet of the file
      currentWorksheet = workbook_in["X STRs"]
      # Creates an empty dictionary to hold all discovered STRs
      Alleles = {}
      # Increment through all rows in the sheet
      for rowNumber in range(1,len(currentWorksheet['A'])):
         # If the riownumber falls within the upper section bounds
         if rowNumber >= xSTRUpperStart and rowNumber <= xSTRUpperStop:
            # If the first cell is not None, then add a new entry to the dictionary
            if (currentWorksheet.cell(row = rowNumber, column=1).value != xr.empty_cell.value and currentWorksheet.cell(row = rowNumber, column=3).value != "INC"):
               Alleles[currentWorksheet.cell(row = rowNumber, column=1).value] = [currentWorksheet.cell(row = rowNumber, column=2).value, currentWorksheet.cell(row = rowNumber, column=3).value]
#               print("Temp Dict entry made for X STR:", currentWorksheet.cell(row = rowNumber, column=1).value,"as", Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value))
         # If the rownumber falls within the lower section bounds, then update the dictionary value for the locus
         if rowNumber >= xSTRLowerStart:
            # Checks if the allele has any reads (does not update anything if there are no reads)
            if (currentWorksheet.cell(row = rowNumber, column=4).value != xr.empty_cell.value and currentWorksheet.cell(row = rowNumber, column=4).value != 0):
#               print ("Sample check:", currentWorksheet.cell(row = rowNumber, column=1).value, sample.sourceDict.get(currentWorksheet.cell(row = rowNumber, column=1).value))
               # Adds the locus
               if (sample.sourceDict.get(currentWorksheet.cell(row = rowNumber, column=1).value) == None):
#                   print ("No locus found yet in sample sourceDict. Checking Allele to begin locus with:", currentWorksheet.cell(row = rowNumber, column=1).value, Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value))
                   sample.sourceDict[currentWorksheet.cell(row = rowNumber, column=1).value] = ["STR", [[currentWorksheet.cell(row = rowNumber, column=2).value,currentWorksheet.cell(row = rowNumber, column=3).value,currentWorksheet.cell(row = rowNumber, column=4).value,currentWorksheet.cell(row = rowNumber, column=5).value, Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[0], [Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[1]]]]]
               # Updates the entry for the locus in the dictionary
               else:
                   oldLocusArray = sample.sourceDict.get(currentWorksheet.cell(row = rowNumber, column=1).value)
#                   print ("Trying to append:", [currentWorksheet.cell(row = rowNumber, column=2).value,currentWorksheet.cell(row = rowNumber, column=3).value,currentWorksheet.cell(row = rowNumber, column=4).value,currentWorksheet.cell(row = rowNumber, column=5).value, Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[0], Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[1]], "to", oldLocusArray)
#                   print ("Test1:", oldLocusArray[0], oldLocusArray[1])
                   oldLocusArray[1].append([currentWorksheet.cell(row = rowNumber, column=2).value,currentWorksheet.cell(row = rowNumber, column=3).value,currentWorksheet.cell(row = rowNumber, column=4).value,currentWorksheet.cell(row = rowNumber, column=5).value, Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[0], [Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[1]]])
#                   print ("Test2:", oldLocusArray[1])
#                   print ("New entry:", newLocusArray)
                   sample.sourceDict[currentWorksheet.cell(row = rowNumber, column=1).value] = oldLocusArray
      
      # Switch to the "Autosomal STRs" sheet of the workbook
      currentWorksheet = workbook_in["Autosomal STRs"]
      # Creates an empty dictionary to hold all discovered STRs
      Alleles = {}
      # Increment through all rows in the sheet
      for rowNumber in range(1,len(currentWorksheet['A'])):
         if rowNumber >= aSTRUpperStart and rowNumber <= aSTRUpperStop:
            # If the first cell is not None, then add a new entry to the dictionary
            if (currentWorksheet.cell(row = rowNumber, column=1).value != xr.empty_cell.value and currentWorksheet.cell(row = rowNumber, column=3).value != "INC"):
               Alleles[currentWorksheet.cell(row = rowNumber, column=1).value] = [currentWorksheet.cell(row = rowNumber, column=2).value, currentWorksheet.cell(row = rowNumber, column=3).value]
#               print("Temp Dict entry made for X STR:", currentWorksheet.cell(row = rowNumber, column=1).value,"as", Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value))
         # If the rownumber falls within the lower section bounds, then update the dictionary value for the locus
         if rowNumber >= aSTRLowerStart:
            # Checks if the allele has any reads (does not update anything if there are no reads)
            if (currentWorksheet.cell(row = rowNumber, column=4).value != xr.empty_cell.value and currentWorksheet.cell(row = rowNumber, column=4).value != 0):
#               #print ("Sample check:", currentWorksheet.cell(row = rowNumber, column=1).value, sample.sourceDict.get(currentWorksheet.cell(row = rowNumber, column=1).value))
               # Adds the locus
               if (sample.sourceDict.get(currentWorksheet.cell(row = rowNumber, column=1).value) == None):
#                   print ("No locus found yet in sample sourceDict. Checking Allele to begin locus with:", currentWorksheet.cell(row = rowNumber, column=1).value, Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value))
                   sample.sourceDict[currentWorksheet.cell(row = rowNumber, column=1).value] = ["STR", [[currentWorksheet.cell(row = rowNumber, column=2).value,currentWorksheet.cell(row = rowNumber, column=3).value,currentWorksheet.cell(row = rowNumber, column=4).value,currentWorksheet.cell(row = rowNumber, column=5).value, Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[0], [Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[1]]]]]
               # Updates the entry for the locus in the dictionary
               else:
                   oldLocusArray = sample.sourceDict.get(currentWorksheet.cell(row = rowNumber, column=1).value)
#                   print ("Trying to append:", [currentWorksheet.cell(row = rowNumber, column=2).value,currentWorksheet.cell(row = rowNumber, column=3).value,currentWorksheet.cell(row = rowNumber, column=4).value,currentWorksheet.cell(row = rowNumber, column=5).value, Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[0], Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[1]], "to", oldLocusArray)
#                   print ("Test1:", oldLocusArray[0], oldLocusArray[1])
                   oldLocusArray[1].append([currentWorksheet.cell(row = rowNumber, column=2).value,currentWorksheet.cell(row = rowNumber, column=3).value,currentWorksheet.cell(row = rowNumber, column=4).value,currentWorksheet.cell(row = rowNumber, column=5).value, Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[0], [Alleles.get(currentWorksheet.cell(row = rowNumber, column=1).value)[1]]])
#                   print ("Test2:", oldLocusArray[1])
#                   print ("New entry:", newLocusArray)
                   sample.sourceDict[currentWorksheet.cell(row = rowNumber, column=1).value] = oldLocusArray
                   
   # Creates a list of locus ids only to use in determining the kit used
   tempLoci = []
   # Goes through all source samples to get locus names
   for key, value in sample.sourceDict.items():
#      print (" XLS Locus:", key, "Value", value)
      tempLoci.append(key)
   # Sets the kit variable using the loci present in the list
   sample.kit = source(tempLoci)[1]
   # Debugging sample print statement
   if (False):
         sample.toString()
   sample.formatDict()
   if (False):
         sample.toString()
   # Returns the sample to the caller
   return sample


# Helper function that identifies a single locus as a SNP or STR
# locus - name of the locus being identified
# Returns a string ('SNP', 'STR', 'LNF') stating what type of locus it is
def  identifyLocusType(locus):
   # If the locus is found in the STR list then return 'STR'
   if locus in LocusHelper.allSTRs:
       return "STR"
   # If the locus is found in the SNP list then return 'SNP'
   elif locus in LocusHelper.allSNPs:
       return "SNP"
   # If the locus was not found in either list, then return 'LNF' (Locus Not Found)
   else:
       print (locus,"Locus not found")
       return "LNF"
 
# Helper function to find the panel that a whole sample was run through
# locuses - a list of loci seen within a sample
# Returns a list that has the Instrument (or company) and kit
def source(locuses):
#   print (locuses)
   # Creates a list of the total sizes of the currently available kits
   kitLocusSizes = [len(LocusHelper.thermoID),
                     len(LocusHelper.miseqID),
                     len(LocusHelper.thermoANC),
                     len(LocusHelper.miseqANC),
                     len(LocusHelper.miseqPHENO),
                     len(LocusHelper.thermoSTRs),
                     len(LocusHelper.miseqSTRs),
                     len(LocusHelper.ySNPs)]
   # Creates a tally table to keep track of how many of each kit are found
   locusCounts = [0,0,0,0,0,0,0,0]
   #Goes through all loci in a sample
   for loc in locuses:
      # If a locus is an identifying locus (e.g., not found in any other kit) then that locus is incremented
       if loc in LocusHelper.thermoID:
#                print("Found in Thermo-ID")
           locusCounts[0] = locusCounts[0] + 1
       elif loc in LocusHelper.miseqID:
#                print("Found in MiSeq-ID")
           locusCounts[1] = locusCounts[1] + 1
       elif loc in LocusHelper.thermoANC:
#                print("Found in Thermo-ANC")
           locusCounts[2] = locusCounts[2] + 1
       elif loc in LocusHelper.miseqANC:
#                print("Found in MiSeq-ANC")
           locusCounts[3] = locusCounts[3] + 1
       elif loc in LocusHelper.miseqPHENO:
#                print("Found in MiSeq-Pheno")
           locusCounts[4] = locusCounts[4] + 1
       elif loc in LocusHelper.thermoSTRs:
#                print("Found in Thermo-STRs")
           locusCounts[5] = locusCounts[5] + 1
       elif loc in LocusHelper.miseqSTRs:
#                print("Found in MiSeq-STRs")
           locusCounts[6] = locusCounts[6] + 1
       elif loc in LocusHelper.ySNPs:
#                print("Found in MiSeq-STRs")
           locusCounts[7] = locusCounts[7] + 1
       # In cases where a locus is found in more than one kit, then it is just skipped
       elif loc in LocusHelper.bothID:
#           print(loc, ",Found in both-ID")
           pass
       elif loc in LocusHelper.bothANC:
#           print(loc,", Found in both-ANC")
           pass
       elif loc in LocusHelper.bothSTRs:
#           print(loc, ", Found in both-STRs")
           pass
       # If the locus is not found in any of the lists, an error message is printed
       else:
           print (loc, "Unknown source......oooohoohohoho")
#   print(locusCounts)
   # Creates an empty list thaty contains the percentage of snps found for a kit
   percentOfSNPsinkit = [0] * len(locusCounts)
   # Goes through each kit to calculate the % of SNPs found
   for k in range(0,len(locusCounts)):
#       print (k, locusCounts[k], kitLocusSizes[k])
       percentOfSNPsinkit[k] = round(100*(locusCounts[k]/kitLocusSizes[k]), 3)
#   print(percentOfSNPsinkit)
   # If the first element of the list is greater than 75% then it is determined to be from a Precision ID Identiy Panel
   if (percentOfSNPsinkit[0] > 75):
       return ["Thermo", "Precision ID"]
   # If the third element of the list is greater than 75% then it is determined to be from a Precision ID Ancestry Panel
   elif (percentOfSNPsinkit[2] > 75):
       return ["Thermo", "Precision Ancestry"]
   # If the sixth element of the list is greater than 75% then it is determined to be from the Thermo STR kit
   elif (percentOfSNPsinkit[5] > 75):
       return ["Thermo", "Thermo STR Kit Name"]
   # If the second element of the list is greater than 75% and the fourth and fifth elements are less than 25% then it is determined to be from the ForenSeq Primer A Panel
   elif (percentOfSNPsinkit[1] > 75 and percentOfSNPsinkit[3] < 25  and percentOfSNPsinkit[4] < 25) or (percentOfSNPsinkit[1] > 75 and percentOfSNPsinkit[3] > 75  and percentOfSNPsinkit[4] > 75):
       return ["MiSeq", "ForenSeq Signature Prep Kit v3.0 Primer A"]
   # If the second element of the list is greater than 75% and the fourth and fifth elements are also greater than 75% then it is determined to be from the ForenSeq Primer B Panel
   elif (percentOfSNPsinkit[1] < 25 and percentOfSNPsinkit[3] > 99  and percentOfSNPsinkit[4] > 25):
       return ["MiSeq", "ForenSeq Signature Prep Kit v3.0 Primer B"]
   # If none of these combinations of percentages are met, then it is from an undetermined kit and an empty list is returned
   elif percentOfSNPsinkit[3] == 0 and max(percentOfSNPsinkit)==percentOfSNPsinkit[4]:
       return ["MiSeq", "ForenSeq Signature Prep Kit v3.0 Primer B"]
   else:
       print ("Unknown kit:", percentOfSNPsinkit)
       return ["Unknown Manufacturer", "Unknown Kit"]

# Helper function to output a list of samples as CSVs according to the boolean command line arguments
# samplelist is the list of samples to output. fileout is the name to output the file as, which will be appended with the current argument.
# All other parameters are booleans to show that information or not
def csvexport(samplelist, showCoverage,showGenotypes,showFlags,showMAF,showSNPs,showSTRs, fileout):
   #print("fo:", fileout)
   # A list of samples that contain SNPs in their loci
   SNPsamplenamelist = []
   # A list of samples that contain STRs in their loci
   STRsamplenamelist = []
   # A list of samples that contain SNPs in their loci
   SNPsamplelist = []
   # A list of samples that contain STRs in their loci
   STRsamplelist = []
   # A list that will contain all SNPs seen throughout all samples
   SNPlocusnamelist = []
   # A list that will contain all STRs seen throughout all samples
   STRlocusnamelist = []
   # A list/array to hold all sample SNP info as displayed in the csv file (for easier manipulation)
   SNPsampleMatrix = []
   # A list/array to hold all sample STR info as displayed in the csv file (for easier manipulation)
   STRsampleMatrix = []
   # Goes through each sample in the list to find all samples that contain SNPs/STRs and all loci found through the list
   # i.e, A master sample and locus list
   for sample in samplelist:
      # Master count number for loci found in all samples
      locuscount = 0
      # Goes through each key and value within the current sample
      for loci,value in sample.finalDict.items():
          # If the current locus is tagged as a SNP
          if value[0] == "SNP":
            # If the current sample is not already in the SNP sample list 
            if sample.name+"("+sample.sampleSource+"-"+sample.kit[1]+")" not in SNPsamplenamelist:
               # Add the sample name to the SNP sample list
#               print ("Pre:", SNPsamplenamelist)
               SNPsamplenamelist.append(sample.name+"("+sample.sampleSource+"-"+sample.kit[1]+")")
               SNPsamplelist.append(sample)
#               print ("Adding:", sample.name+"("+sample.sampleSource+"-"+sample.kit[1]+")")
#               print ("Post:", SNPsamplenamelist)
            # If the locuscount has not had any samples added yet
            elif locuscount == 0:
#               print("Sample already in list")
               # Add the first entry to the samplenameList
               SNPsamplenamelist.append(sample.name+"("+sample.sampleSource+"-"+sample.kit[1]+")")
               SNPsamplelist.append(sample)
            # if the loci doesnt exist in the loci name list yet
            if loci not in SNPlocusnamelist:
               # Add the locus name to the SNP locus list
#               print("!: loci not found:",loci)
               SNPlocusnamelist.append(loci)
          # If the current locus is tagged as a STR
          if value[0] == "STR":
            # If the current sample is not already found in the STR sample list
            if sample.name+"("+sample.sampleSource+"-"+sample.kit[1]+")" not in STRsamplenamelist:
               # Add the sample to the STR sample list
               STRsamplenamelist.append(sample.name+"("+sample.sampleSource+"-"+sample.kit[1]+")")
               STRsamplelist.append(sample)
            # If no locuses have been added to the list yet
            elif locuscount == 0:
#               print("Sample already in list")
               # Adds sample name to the list
               SNPsamplenamelist.append(sample.name+"("+sample.sampleSource+"-"+sample.kit[1]+")")
               SNPsamplelist.append(sample)
#            print ("v",loci,value)
            # If locus is not already in the STR locus list
            if loci not in STRlocusnamelist:
               # Add locus name to the STR locus list
               STRlocusnamelist.append(loci)
#            for alleles in value[1]:
#               print("a:",sample.name,loci,alleles[1])
          locuscount = locuscount + 1
    # Used in debugging to check lists
   if (False):
      print("SNP sample list:",SNPsamplenamelist)
      print("STR sample list:",STRsamplenamelist)
      print("SNP locus list ("+str(len(SNPlocusnamelist))+"):",sorted(SNPlocusnamelist))
      print("STR locus list ("+str(len(STRlocusnamelist))+"):",sorted(STRlocusnamelist))
   
   # If the parameter for STRs is marked in the command line to create STR report sheet
   if (showSTRs):
#      print ("Doing STRs")
      # Creates a variable to hold the fule header line number of samples * number of traits requested
      STRsamplenamelistwithCopies = []
      headerLine = [" "," "," "," "," "]
      # Extends the header line by the length of the number of samples found, if genotpes are requested
      if showGenotypes == True:
         STRsamplenamelistwithCopies = deepcopy(STRsamplenamelist)
         STRsamplenamelistwithCopies.append(" ")
         headerLine.append("Genotypes")
         for i in range(0,len(STRsamplenamelist)):
            headerLine.append(" ")
      # Extends the header line by the length of the number of samples found, if flags are requested      
      if showFlags == True:
         if STRsamplenamelistwithCopies == None:
            STRsamplenamelistwithCopies = deepcopy(STRsamplenamelist)
         else:
            for samplename in STRsamplenamelist:
               STRsamplenamelistwithCopies.append(samplename)
            STRsamplenamelistwithCopies.append(" ")
         headerLine.append("Flags")
         for i in range(0,len(STRsamplenamelist)):
            headerLine.append(" ")
      # Extends the header line by the length of the number of samples found, if coverage is requested      
      if showCoverage == True:
         if STRsamplenamelistwithCopies == None:
            STRsamplenamelistwithCopies = deepcopy(STRsamplenamelist)
         else:            
            for samplename in STRsamplenamelist:
               STRsamplenamelistwithCopies.append(samplename)
            STRsamplenamelistwithCopies.append(" ")
         headerLine.append("Coverages")
         for i in range(0,len(STRsamplenamelist)):
            headerLine.append(" ")
      # Extends the header line by the length of the number of samples found, if MAFs are requested
      if showMAF == True:
         if STRsamplenamelistwithCopies == None:
            STRsamplenamelistwithCopies = deepcopy(STRsamplenamelist)
         else:
            for samplename in STRsamplenamelist:
               STRsamplenamelistwithCopies.append(samplename)
            STRsamplenamelistwithCopies.append(" ")
         headerLine.append("MAFs")
         for i in range(0,len(STRsamplenamelist)):
            headerLine.append(" ")
#      print (headerLine)
     
      # Creates the new filename to save to for genotypes
      ofile  = open(fileout+'_STR.csv', "w", newline="\n", encoding="utf-8")
      # Creates the writer object to use to save genotype information
      writer = csv.writer(ofile)
      # Makes a deepcopy of the headerline, so as not to screw up references, Is also an accessible matrix that contains all printed values as shown
      STRsampleMatrix = deepcopy([headerLine])
#      writer.writerow(headerLine)
      # Creates the second header line with proper headers and all the samples at the correct number of copies for the requested report
      STRsampleMatrix.append(['STR Locus','Chromosome','Source Kit', 'Allele'," "]+STRsamplenamelistwithCopies)
#      writer.writerow(['STR Locus','Chromosome','Source Kit', 'Allele', " "]+STRsamplenamelistwithCopies)
      # Goes through all loci that have been found in the samples
      for locus in STRlocusnamelist:
           # Creates an empty list of values that exist at the current locus
           #The values for the row to be written, Each locus could have multiple alleles and until that is known a list  is used to hold all alleles
           rowsToWrite = []
####           # ???
####           locilist = []
           # A variable to hold all the different alleles for a certain locus
           uniqueAlleleList = []
           # A variable to hold the number of times each allele appears in the list
           uniqueAlleleCounts = []
           #Variables to keep track of starting position in the lists for data requested (genotypes, etc...)
           shiftValue = len(STRsamplenamelist)+1
           # Start position of Genotypes
           genotypeIndex = 5
           # Start position of flags according to number of samples
           flagIndex = genotypeIndex + shiftValue
           # Start position of coverages according to number of samples
           coverageIndex = flagIndex + shiftValue
           # Start position of MAFs according to number of samples
           MAFIndex = coverageIndex + shiftValue
#           print (genotypeIndex, flagIndex, coverageIndex, MAFIndex)
            # Keeps all start positions for the attributes in a list
           Indices = [genotypeIndex, flagIndex, coverageIndex, MAFIndex]
           # Goes through each sample found in the STR list
           for sample in STRsamplelist:
#               print ("Sample",sample.name, " | Locus:", locus, "Type:", type(locus))
#              sample.toString()
               # If the current locus is found in the current sample process it for unique alleles
               if locus in sample.finalDict:
                  # If any of the alleles at the current locus in the current sample have not been added to the unique alleles list, then do so
                   for allele in sample.getAlleles(locus):
#                      print ("Allele:",  type(allele[0]))
                      if (float(allele[0]) not in uniqueAlleleList):
                           uniqueAlleleList.append(float(allele[0]))
####                   locilist.append([sample.name,sample.getAlleles(locus)])
               # Sort unique alleles as floats to sort by numerical value instead of string
               uniqueAlleleList = sorted(uniqueAlleleList)
#               print ("Pre Pre Unique Allele List", uniqueAlleleList, sample.name, locus)
           # Goes through each unique allele in the list to find and edit to be a corrected edited string value
           for i in range(0, len(uniqueAlleleList)):
#               print ("Pre Unique Allele List", uniqueAlleleList[i], type(uniqueAlleleList[i]))
               if (str(uniqueAlleleList[i]).find(".0")) > -1:
                  uniqueAlleleList[i] = str(uniqueAlleleList[i])[:-2]
               else:
                  uniqueAlleleList[i] = str(uniqueAlleleList[i])
#               print ("Post Unique Allele List", uniqueAlleleList[i], type(uniqueAlleleList[i]))
           # Goes through each allele in the list...again, after it been formatted to a string
           # Will find the largest occuring allele to make sure there is room when processing the samples with fewer alleles
           # Creates the alleles with counts variable that pairs alleles with their number of occurences
           for uniqueAllele in uniqueAlleleList:
#               print ("Unique Allele:", uniqueAllele, "Type?", type(uniqueAllele))
               largestAlleleCount = 0
               for sample in samplelist:
#                    print ("ZAZZY:", sample.getSpecificAllele(locus, uniqueAllele))
                    if (sample.getSpecificAllele(locus, uniqueAllele) != None):
                        currentSampleAlleleCount = len(sample.getSpecificAllele(locus, uniqueAllele))
                        if currentSampleAlleleCount > largestAlleleCount:
                            largestAlleleCount = currentSampleAlleleCount
               uniqueAlleleCounts.append([uniqueAllele, largestAlleleCount])
#               print (locus, "Loci List", locilist)
#               print ("Unique Allele Count", uniqueAlleleCounts)
           # Goes through each tuple of alleleswithcounts
           for allelewithcount in uniqueAlleleCounts:
               # Creates a 0
               currentCount = 0
               # Creates an empty list to hold each row for an allele
               alleleRowstoWrite = []
#               print ("Testing if allelewithcount has a length:",allelewithcount)
#               print ("Next:", allelewithcount[0], "Locus:", locus)
               # Appends the template for a row a cretain number of times  based on thenumber of alleles found
               for i in range(0, allelewithcount[1]):
#                  print(locus)
#                  print(LocusHelper.locusPositions.get(locus)[0])
#                  print (getKit(locus))
#                  print (allelewithcount[0])
                  alleleRowstoWrite.append([locus, LocusHelper.locusPositions.get(locus)[0], getKit(locus), allelewithcount[0], " "])
#               print ("X",alleleRowstoWrite)
               # Creates an empty dictionary to help track the total number of rows needed
               samplesAllelesDictforCurrentLoci = {}
               # Goes through all smaples in the STR sample name list
               for sample in STRsamplelist:
#                  sample.toString()
                  # Checks if the current sample has any alleles of the current locus
                  if sample.getSpecificAllele(locus, allelewithcount[0]) != None:
#                     print ("OOO",sample.getSpecificAllele(locus, allelewithcount[0]))
                     samplesAllelesDictforCurrentLoci.update({sample.name:sample.getSpecificAllele(locus, allelewithcount[0])})
#               for key, value in samplesAllelesDictforCurrentLoci.items():
#                  print("Key", key, "Value", value)
#               print ("Sample size:", len(STRsamplelist))
               # Adds the Genotype info to be added to the sample matrix
               # Goes through all smaples in the STR sample name list to get genotypes
               for sample in STRsamplelist:
                  # Goes through to check and add values based on the max number of occurences of the current allele
                  for i in range(0, allelewithcount[1]):
#                     print ("Before:", alleleRowstoWrite)
#                     print ("i:", i)
#                     print("SADCL for", sample.name," and loci;", locus,":", samplesAllelesDictforCurrentLoci.get(sample.name))
                     # If this sample locus / # has a(n) allele
                     if (samplesAllelesDictforCurrentLoci.get(sample.name)) != None:
#                        print("Looking for locus:", locus, "allele:", allelewithcount[0], "in",sample.name, "with a total count of", allelewithcount[1], "and local size of", len(samplesAllelesDictforCurrentLoci.get(sample.name)))
                        # If the sample locus has more than the current number of alleles being processed print the # allele value for genotype
                        if (i < len(samplesAllelesDictforCurrentLoci.get(sample.name))):
                           alleleRowstoWrite[i].append(str(samplesAllelesDictforCurrentLoci.get(sample.name)[i][2]))
                        #No other variants of this allele were found
                        else:
                           alleleRowstoWrite[i].append("-") 
#                        print("Z:", alleleRowstoWrite)
                     # No allelle was found for this sample locus / #
                     else:
#                        print("Locus:", locus, "allele:", allelewithcount[0], "in",sample.name, "with a total count of", allelewithcount[1], "and local size of NONE")
                        alleleRowstoWrite[i].append("-")
#                     print ("After:", alleleRowstoWrite)
#               print ("A:", alleleRowstoWrite)
               # Goes through each row and adds an empty space for a column buffer between attributes
               for row in alleleRowstoWrite:
                  row.append(" ")
               # Goes through all smaples in the STR sample name list to get genotypes
               for sample in STRsamplelist:
                  # Goes through to check and add values based on the max number of occurences of the current allele
                  for i in range(0, allelewithcount[1]):
#                     print ("Before:", alleleRowstoWrite)
#                     print ("i:", i)
                     # If this sample locus / # has a(n) allele
                     if (samplesAllelesDictforCurrentLoci.get(sample.name)) != None:
#                        print("Looking for locus:", locus, "allele:", allelewithcount[0], "in",sample.name, "with a total count of", allelewithcount[1], "and local size of", len(samplesAllelesDictforCurrentLoci.get(sample.name)))
                        # If the sample locus has more than the current number of alleles being processed print the # allele value for flags
                        if (i < len(samplesAllelesDictforCurrentLoci.get(sample.name))):
#                           print ("Appending:", samplesAllelesDictforCurrentLoci.get(sample.name)[i][5], "-")
                           # Checks to see if Flags value is an empty string and replaces it with a definitive statement of no flags
                           if (samplesAllelesDictforCurrentLoci.get(sample.name)[i][5] == ""):
                              alleleRowstoWrite[i].append("No Flags")
                           # Otherwise the value of the flag is used as is
                           else:
                              alleleRowstoWrite[i].append(str(samplesAllelesDictforCurrentLoci.get(sample.name)[i][5]))
                        #No other variants of this allele were found
                        else:
                           alleleRowstoWrite[i].append("-") #No other variants of this allele were found
#                        print("Z:", alleleRowstoWrite)
                     # No allelle was found for this sample locus / #
                     else:
#                        print("Locus:", locus, "allele:", allelewithcount[0], "in",sample.name, "with a total count of", allelewithcount[1], "and local size of NONE")
                        alleleRowstoWrite[i].append("-")
#                     print ("After:", alleleRowstoWrite)
#               print ("B:", alleleRowstoWrite)
               # Goes through each row and adds an empty space for a column buffer between attributes
               for row in alleleRowstoWrite:
                  row.append(" ")
               # Goes through all smaples in the STR sample name list to get genotypes
               for sample in STRsamplelist:
                  # Goes through to check and add values based on the max number of occurences of the current allele
                  for i in range(0, allelewithcount[1]):
#                     print ("Before:", alleleRowstoWrite)
#                     print ("i:", i)
                     # If this sample locus / # has a(n) allele                          
                     if (samplesAllelesDictforCurrentLoci.get(sample.name)) != None:
#                        print("Looking for locus:", locus, "allele:", allelewithcount[0], "in",sample.name, "with a total count of", allelewithcount[1], "and local size of", len(samplesAllelesDictforCurrentLoci.get(sample.name)))
                        # If the sample locus has more than the current number of alleles being processed print the # allele value for coverages 
                        if (i < len(samplesAllelesDictforCurrentLoci.get(sample.name))):
#                           print ("Appending:", samplesAllelesDictforCurrentLoci.get(sample.name)[i][1])
                           alleleRowstoWrite[i].append(str(samplesAllelesDictforCurrentLoci.get(sample.name)[i][1]))
                        # No other variants of this allele were found
                        else:
                           alleleRowstoWrite[i].append("-") #No other variants of this allele were found
#                        print("Z:", alleleRowstoWrite)
                     # No allelle was found for this sample locus / #
                     else:
#                        print("Locus:", locus, "allele:", allelewithcount[0], "in",sample.name, "with a total count of", allelewithcount[1], "and local size of NONE")
                        alleleRowstoWrite[i].append("-")
#                     print ("After:", alleleRowstoWrite)
#               print ("C:", alleleRowstoWrite)
               # Goes through each row and adds an empty space for a column buffer between attributes
               for row in alleleRowstoWrite:
                  row.append(" ")
               #Adds in the MAF to the STR matrix sheet
               for sample in STRsamplelist:
                  for i in range(0, allelewithcount[1]):
#                     print ("Before:", alleleRowstoWrite)
#                     print ("i:", i)
                     # If this sample locus / # has a(n) allele      
                     if (samplesAllelesDictforCurrentLoci.get(sample.name)) != None:
#                        print("Looking for locus:", locus, "allele:", allelewithcount[0], "in",sample.name, "with a total count of", allelewithcount[1], "and local size of", len(samplesAllelesDictforCurrentLoci.get(sample.name)))
                        # If the sample locus has more than the current number of alleles being processed print the # allele value for MAFs
                        if (i < len(samplesAllelesDictforCurrentLoci.get(sample.name))):
#                           print ("Appending:", samplesAllelesDictforCurrentLoci.get(sample.name)[i][1])
                           alleleRowstoWrite[i].append(str(samplesAllelesDictforCurrentLoci.get(sample.name)[i][4]))
                        # No other variants of this allele were found
                        else:
                           alleleRowstoWrite[i].append("-") #No other variants of this allele were found
#                        print("Z:", alleleRowstoWrite)
                     # No allelle was found for this sample locus / #
                     else:
#                        print("Locus:", locus, "allele:", allelewithcount[0], "in",sample.name, "with a total count of", allelewithcount[1], "and local size of NONE")
                        alleleRowstoWrite[i].append("-")
#                     print ("After:", alleleRowstoWrite)
#               print ("D:", alleleRowstoWrite)
               # Adds each row of in Alleesrowstowrite to the full STRmatrix
               for row in alleleRowstoWrite:
                  STRsampleMatrix.append(row)
#           print (STRsampleMatrix)
      # Goes through all rows in STRmatrix to print to file
      for row in STRsampleMatrix:
#         print(row)
         # If it is the first couple rows with little content, just print the row (headers and such)
         if row[0] == "" or row[0] == " " or row[0] == "STR Locus":
#            print("Whole row print:", row)
            writer.writerow(row)
         # Proceed on to content rows and creates an edited row based on what the user requested on the command line
         else:
            # Starts with the locus id, chromosome, source kit, and allele
            editedRow = row[:5]
            # Adds the genotype information if requested
            if showGenotypes == True:
#               print ("Genotypes:", row[Indices[0]:Indices[1]])
               editedRow += row[Indices[0]:Indices[1]]
            # Adds the flag information if requested
            if showFlags == True:
#               print ("Flags:", row[Indices[1]:Indices[2]])
               editedRow += row[Indices[1]:Indices[2]]
            # Adds the coverage information if requested
            if showCoverage == True:
#               print ("Coverages:", row[Indices[2]:Indices[3]], "Indices:", Indices[2], Indices[3])
               editedRow += row[Indices[2]:Indices[3]]
            # Adds the MAF information if requested
            if showMAF == True:
#               print ("MAFs:", row[Indices[3]:])
               editedRow += row[Indices[3]:]
            # Writes the line to file
            writer.writerow(editedRow)
      # Closes the STR file writer
      ofile.close()      
                 
   # If the showSNPs boolean is set (The user wants to see details about SNP loci)      
   if (showSNPs):
         # Creates the new filename to save to for genotypes
         ofile  = open(fileout+'_SNPs.csv', "w", newline="\n", encoding="utf-8")
         # Creates the writer object to use to save genotype information
         writer = csv.writer(ofile)
         # Creates a holder for the full list of names with the number of copies needed to head all attribute columns
         SNPsamplenamelistwithCopies = []
         # Creates the template for the header line
         headerLine = [" "," ", " ", " ", " "]
         # Extends the header line by the length of the number of samples found, if genotpes are requested
         if showGenotypes == True:
            SNPsamplenamelistwithCopies = deepcopy(SNPsamplenamelist)
            headerLine.append("Genotypes")
            for i in range(0,len(SNPsamplenamelist)):
               headerLine.append(" ")
         # Extends the header line by the length of the number of samples found, if flags are requested
         if showFlags == True:
            if SNPsamplenamelistwithCopies == None:
               SNPsamplenamelistwithCopies = deepcopy(SNPsamplenamelist)
            else:
               SNPsamplenamelistwithCopies.append(" ")
               for samplename in SNPsamplenamelist:
                  SNPsamplenamelistwithCopies.append(samplename)
            headerLine.append("Flags")
            for i in range(0,len(SNPsamplenamelist)):
               headerLine.append(" ")
         # Extends the header line by the length of the number of samples found, if coverages are requested
         if showCoverage == True:
            if SNPsamplenamelistwithCopies == None:
               SNPsamplenamelistwithCopies = deepcopy(SNPsamplenamelist)
            else:
               SNPsamplenamelistwithCopies.append(" ")
               for samplename in SNPsamplenamelist:
                  SNPsamplenamelistwithCopies.append(samplename)
            headerLine.append("Coverages")
            for i in range(0,len(SNPsamplenamelist)):
               headerLine.append(" ")
         # Extends the header line by the length of the number of samples found, if MAFs are requested
         if showMAF == True:
            if SNPsamplenamelistwithCopies == None:
               SNPsamplenamelistwithCopies = deepcopy(SNPsamplenamelist)
            else:
               SNPsamplenamelistwithCopies.append(" ")
               for samplename in SNPsamplenamelist:
                  SNPsamplenamelistwithCopies.append(samplename)
            headerLine.append("MAFs")
            for i in range(0,len(SNPsamplenamelist)):
               headerLine.append(" ")
            
         
         # Writes the first row which includes two blanks plus all names of samples that are going to be included
         SNPsampleMatrix = deepcopy([headerLine])
         # Writes the header line to file
         writer.writerow(headerLine)
         # Writes the second header line to file and then to the matrix
         SNPsampleMatrix.append(['SNP Locus','Chromosome', 'Position','Source Kit', " "]+SNPsamplenamelistwithCopies)
         writer.writerow(['SNP Locus','Chromosome', 'Position','Source Kit', " "]+SNPsamplenamelistwithCopies)
#         print (SNPsampleMatrix)
         
#         for locus in SNPlocusnamelist:
#            print(locus)
         # Goes through each locus in the SNP locus list
         for locus in SNPlocusnamelist:
#            print ("##", len(SNPlocusnamelist), "##")
#            print ("@@", locus, "@@")
            # Creates the empty list of values that will be written on a line during export
            locusvalues = [" "]
            # Used to capture all values for a single sample to be used when rearranging list horizontally
            samplevalues = []
            # Goes through each sample in the samplelist 
            for sample in samplelist:
#                print(sample.name, locus)
                # Checks if the current sample is in the SNP sample list
                if sample.name+"("+sample.sampleSource+"-"+sample.kit[1]+")" in SNPsamplenamelist:
                     # If the sample has a value for the current locus
                     if locus in sample.finalDict:
                        tempvals = sample.finalDict.get(locus)[1:]
                        tempvals.insert(0,sample.name)
                        tempvals.insert(1, locus)
                        samplevalues.append(tempvals)
                     # Else adds empty values to the sample value
                     else:
                        samplevalues.append([samplename, locus, "-","-","-","-"])
#            print (samplevalues)
            # If the command line user requested genotypes to be shown
            if showGenotypes == True:
               # For each sample print its genotype 
               for i in range(0, len(samplevalues)):
                   locusvalues.append(samplevalues[i][2])
               # Add a row delimiter ( a space)
               locusvalues.append(" ")
#            print ("1",locusvalues)
            # If the command line user requested genotypes to be shown
            if showFlags == True:
               # For each sample process its flags
                for i in range(0, len(samplevalues)):
           #        print (i)
           #        print (i[0])
           #        print (i[1])
           #        print (i[2])
           #        print (i[3])
                   # If the flag value is not (empty and a float) but its greater than 0
                   if not (isinstance(samplevalues[i][5], float) and samplevalues[i][5] != '') and len((samplevalues[i][5])) > 0:
                       # If the flag is an empty string then report No Flags
                       if samplevalues[i][5] == ['']:
                           locusvalues.append('No Flags')
                       # Else report the given value
                       else:
                           locusvalues.append(samplevalues[i][5])
                   # Otherwise no flags are reported
                   else:
                       locusvalues.append('No Flags')
                # Add a row delimiter ( a space)
                locusvalues.append(" ")
#            print ("2",locusvalues)
            # If the command line user requested coverage to be shown
            if showCoverage == True:
                # For each sample i the list, append the coverage value
                for i in range(0, len(samplevalues)):
                   locusvalues.append(samplevalues[i][3])
                locusvalues.append(" ")
#            print ("3",locusvalues)
            # If the command line user requested MAFs to be shown
            if showMAF == True:
               # For each sample, append the MAF value
                for i in range(0, len(samplevalues)):
                   locusvalues.append(samplevalues[i][4])
                # Add a row delimiter ( a space)
                locusvalues.append(" ")
#            print ("4",locusvalues)              
#            print (len(locusvalues))
#            print("!!",locus, "!!", sample.name)
#            print(LocusHelper.locusPositions.get(locus)[0]+"-")
#            print(LocusHelper.locusPositions.get(locus)[1]+"--"+getKit(locus), locusvalues)
            # Writes the current locus along with kit origin and all values to the next line of the file
            writer.writerow([locus, LocusHelper.locusPositions.get(locus)[0], LocusHelper.locusPositions.get(locus)[1], getKit(locus)]+locusvalues)
            SNPsampleMatrix.append([locus, LocusHelper.locusPositions.get(locus)[0], LocusHelper.locusPositions.get(locus)[1], getKit(locus)]+locusvalues)
            
         # Closes the writer 
         ofile.close()
#         print (SNPsampleMatrix)
          
         #Test area for snp sample matrix
#         ofile2  = open(fileout+'_SNPs.csv', "w", newline="\n", encoding="utf-8")
         # Creates the writer object to use to save genotype information
#         writer = csv.writer(ofile2)
         
#         for line in SNPsampleMatrix:
#            writer.writerow(line)
#         ofile2.close()

# Finds the source kit of a single locus
# locus - the locus to find the source of
# returns a string that contains the name of the kit (or source) that the locus came from
def getKit(locus):
        if locus in LocusHelper.thermoID:
            return "Thermo Precision ID Panel"
        elif locus in LocusHelper.ySNPs:
            return "Thermo Precision ID Panel (Y SNP)" 
        elif locus in LocusHelper.thermoANC:
            return "Thermo Precision Ancestry Panel"
        elif locus in LocusHelper.thermoSTRs:
            return "Thermo STR kit"
        elif locus in LocusHelper.miseqID or locus in LocusHelper.miseqSTRs:
            return "ForenSeq Signature prep Kit Primer A"
        elif locus in LocusHelper.miseqANC or locus in LocusHelper.miseqPHENO:
            return "ForenSeq Signature prep Kit Primer B"
        elif locus in LocusHelper.bothANC:
            return "Cross Platform Ancestry SNP"
        elif locus in LocusHelper.bothID:
            return "Cross Platform Identity SNP"
        elif locus in LocusHelper.bothSTRs:
            return "Cross Platform STR"
        # Returns the 'error' string stating that a match was not found
        else:
            print ("Locus", locus, "is not found in any kit list")
         

 # Standard start of a python program       
if __name__ == "__main__":
   main(sys.argv[1:])
    

        
        

   