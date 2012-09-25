#!/usr/bin/env python


def main():
	f = open("./nmon","r")
	headers={}
	data={}
	header=False
	start=False
	bbbp=False
	for l in f.readlines():
		#catch when we are done with the header
		if bbbp==True and "ZZZZ," in l:
			bbbp=False
			header=False
			start=True
		# start parsing
		if start:
			if "ZZZZ," not in l:
				#print l.strip()
				bits=l.strip().split(',')
				tmp=[]
				if bits[0] in headers.keys():
					tmp=headers[bits[0]]
				
			
		elif header == False:
			if "AAA," not in l:
				header=True
		
		if header == True:
			if "BBBP," in l:
				bbbp=True
			elif bbbp==False:
				bits=l.strip().split(',')
				headers[bits[0]]=bits[1:]
				print bits[0], headers[bits[0]]
			
	

if __name__ == "__main__":
	main()
