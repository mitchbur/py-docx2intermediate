# -*- coding: utf-8 -*-

import zipfile
import xml.etree.ElementTree as etree
import sys
import re

ns_map = { 'wp': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' }

# ignored paragraph styles
ignored_pp_styles = set( [ 'Normal', 'NormalWeb' ] )

for nskey in ns_map.keys():
	etree.register_namespace( nskey, ns_map[nskey] )


#===============================================================================
# TablePos
#   track table row and column positions
#===============================================================================
class TablePos:
	def __init__( self ):
		self.row = 0
		self.col = 0

	def tr_start( self ):
		self.row += 1
		self.col = 0
		return( "<tr row='{0}'>\n".format( self.row ) )

	def tr_end( self ):
		return( '</tr>\n' )

	def tc_start( self ):
		self.col += 1
		return( "<tc col='{0}'>\n".format( self.col ) )

	def tc_end( self ):
		return( '</tc>\n' )

	def tc_span( self, val ):
		self.col += ( val - 1 )

# class to extract the XML bytes from the
# Word document XML member of the .docx compressed archive
class DocxFile:

	def __init__( self, fpath ):
		self.zf = zipfile.ZipFile( fpath )
		self.docfile = None

	def __del__( self ):
		self.close( )

	def close( self ):
		if ( not self.docfile is None ):
			self.docfile.close( )
			self.docfile = None
		if ( not self.zf is None ):
			self.zf.close( )
			self.zf = None

	def open_docxml( self ):
		if ( self.docfile is None and \
		     not self.zf is None ):
			self.docfile = self.zf.open( 'word/document.xml' )
		return self.docfile

def normalized_tag( tg ):
	fq_tag = tg
	fq_match = re.match( "^{([^}]+)}(.*)$", tg )
	if ( not fq_match ):
		pfx_match = re.match( "^([^:]+):(.*)$", tg )
		if ( pfx_match ):
			fq_tag = "{" + ns_map[pfx_match.group(1)] + "}" + pfx_match.group(2)
	return fq_tag

def prefixed_tag( tag ):
	pfx_tag = None
	fq_match = re.match( "^{([^}]+)}(.*)$", tag )
	if ( fq_match ):
		for pfx in ns_map.keys():
			if ( ns_map[pfx] == fq_match.group(1) ):
				pfx_tag = pfx + ":" + fq_match.group(2)
				break
		if ( pfx_tag is None ):
			pfx_tag = ':' + fq_match.group(2)
	else:
		pfx_tag = tag
	return pfx_tag

def transform( docx_fpath, intm_fpath ):
	docxf = DocxFile( docx_fpath )
	with docxf.open_docxml( ) as docx_readio, \
		open( intm_fpath, "w", encoding="utf-8" ) as intm:

		need_cr = False
		tbl_stack = []
		for (event,elem) in etree.iterparse( docx_readio, events=["start","end"] ):
			pfx_tag = prefixed_tag( elem.tag )
			if ( 'wp:t' == pfx_tag ) and ( event == "end" ):
				t_text = elem.text
				intm.write( t_text if t_text is not None else '' )
				need_cr = True
			elif ( 'wp:p' == pfx_tag ) and ( event == "start" ):
				if ( need_cr ):
					intm.write( "\n" )
				intm.write( '<p>' )
				need_cr = True
			elif ( 'wp:pStyle' == pfx_tag ) and ( event == "start" ):
				if ( len ( elem.attrib ) > 0 ):
					style_val = elem.attrib.get( normalized_tag('wp:val'), None )
					if ( not style_val in ignored_pp_styles ):
						intm.write( "<div class='" + style_val + "'/>" )
			elif ( 'wp:tbl' == pfx_tag ):
				if ( need_cr ):
					intm.write( "\n" )
				if ( event == "start" ):
					intm.write( '<table>\n' )
					tbl_stack.append( TablePos() )
				else:
					intm.write( '</table>\n' )
					tbl_stack.pop( )
				need_cr = False
			elif ( 'wp:tr' == pfx_tag ):
				if ( need_cr ):
					intm.write( "\n" )
				if ( event == "start" ):
					intm.write( tbl_stack[len(tbl_stack)-1].tr_start( ) )
				else:
					intm.write( tbl_stack[len(tbl_stack)-1].tr_end( ) )
				need_cr = False
			elif ( 'wp:tc' == pfx_tag ):
				if ( need_cr ):
					intm.write( "\n" )
				if ( event == "start" ):
					intm.write( tbl_stack[len(tbl_stack)-1].tc_start( ) )
				else:
					intm.write( tbl_stack[len(tbl_stack)-1].tc_end( ) )
				need_cr = False
			elif ( 'wp:gridSpan' == pfx_tag ) and ( event == "start" ):
				if ( len ( elem.attrib ) > 0 ):
					span_val = elem.attrib.get( normalized_tag('wp:val'), None )
					tbl_stack[len(tbl_stack)-1].tc_span( int( span_val ) )
			elif ( 'wp:tab' == pfx_tag ) and ( event == "start" ):
				intm.write( '\\t' )
				need_cr = True
			elif ( 'wp:br' == pfx_tag ) and ( event == "start" ):
				intm.write( '<br/>' )
				need_cr = True
			elif ( 'wp:cr' == pfx_tag ) and ( event == "start" ):
				intm.write( '\\r' )
				need_cr = True

		intm.close( )
		docx_readio.close( )

if __name__ == '__main__':
	text = transform( sys.argv[1], sys.argv[2] )
