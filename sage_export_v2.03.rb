=begin
	
	Some link
		https://groups.google.com/forum/#!topic/sketchupruby/cBEOY96g7I0 => help with excel ole properties


	VERSION
	2.0 	+ (améliorations)
			+ Premier essai pour prendre en compte un fichier de conf
			- (correction de bug)
	2.01    + ratio
			+ regroupements
	2.02	+ mouvements
	2.03	+ type de ligne : vide
			+ ajout d'un test pour vérifier les dates de la base de données afin de garder des requetes cohérentes

	TODO
		Add ability to send a list of count to cpta.l functions 

	EXPLANATION
	-d1/1/2013 compute all data between 1/1/2013 and 31/12/2013 (border included)

	sage_export_v2.03.rb -v -d1/1/2013 -l14 (verbose + 14 months)

	DATA
		7% => compte de classe 7 le % est obligatoire
=end 

require "rubygems"
require "sequel"
require "awesome_print"
require 'Date'
#
require 'logger'

require 'optparse'

options = {}
OptionParser.new do |opts|
  opts.banner = "Usage: sage_export.rb [options]"

  opts.on("-b", "--build", "Build database") do |v|
    options[:build] = true
  end  

  opts.on("-d dd/mm/yyyy", "--date dd/mm/yyyy", "Date to calculate the value of a cell (better to be on first day month of year) default is today") do |v|
  	options[:date] = v
  end

  opts.on("-l 12, --length 12", "length  in month (12 is default)") do |v|
    options[:length] = v
  end 

  opts.on("-s filename.sql, --src filename.sql", "Source of SQL data") do |v|
    options[:sqlsrc] = v
  end

   opts.on("-v", "--[no-]verbose", "Run verbosely") do |v|
    options[:verbose] = v
  end 

end.parse!


# default parameters

options[:sqlsrc] ||='sage.sql'
options[:build] ||= false
options[:date]  ||= Date.today().strftime("%d/%m/%Y")
options[:verbose] ||= false
options[:length] ||= 12


# quelques exemples de requetes qui fonctionnent
=begin
	
SELECT sum("EC_MONTANT") AS "SUM" FROM "F_ECRITUREC" WHERE (( CG_NUM LIKE '6063%' OR  CG_NUM LIKE '6064%'  and EC_SENS=1) AND ("JM_DATE" >= {d '2013-12-01'}) AND ("JM_DATE" < {d '2014-01-01'}))

SELECT sum(ec_montant) FROM "F_ECRITUREC" WHERE (( CG_NUM LIKE '6063%' OR  CG_NUM LIKE '6064%') and ec_sens=0 and jm_date = '2013-09-01')

 
=end 


class Numeric
  Alph = ("A".."Z").to_a
  def alph
    s, q = "", self
    (q, r = (q - 1).divmod(26)) && s.prepend(Alph[r]) until q.zero?
    s
  end
end


#TEST
#Ruby 1.9 has removed the current directory from the load path, and so you will need to do a relative require on this file, as Pascal says: require_relative or ./ before filename

require "./lib/compta.rb"

GC.start

sFileName = "BeWe_Compte-exploitation-realise"
sFormatComptabilite = '_-* # ##0,00 €_-;-* # ##0,00 €_-;_-* "-"?? €_-;_-@_-'
sFormatPourcentage = '0,00%;[Rouge] -0,00%'


File.delete("#{sFileName}.xls") if File.exists?("#{sFileName}.xls")
File.delete("./db/sage.db") if File.exists?("./db/sage.db") && options[:build]


#DB = Sequel.odbc('compta_32', :user => "<Administrateur>", :password => "", :db_type=> 'progress', :loggers => [Logger.new($stdout)])
DB1= Sequel.connect('sqlite://db/sage.db', :logger => Logger.new('log/db.log'))
if (options[:build])
	puts "building database"
	DB1.run File.read("./#{options[:sqlsrc]}")
	DB1.run File.read("./sql/F_ECRITUREC.sql")
else
	puts "using old database"
end


#DB.loggers << Logger.new($stdout)
#dbEc = DB[:F_ECRITUREC]
#ap dbEc.select(:ec_montant).sum(:ec_montant)
#ap dbEc.select(:ec_montant).where("CG_NUM LIKE '7%'").sum(:ec_montant)
#ap dbEc.select(:ec_montant).where("CG_NUM LIKE '7%'").filter{ (jm_date >= f>>m) }.sum(:ec_montant) 
#ap dbEc.select(:ec_montant).where("CG_NUM LIKE '7%'").filter{ (jm_date >= f>>m)  & (jm_date < f>> (m +1 )) }.sum(:ec_montant) 
#ap dbEc.select(:ec_montant).where("CG_NUM LIKE '61%' or CG_NUM LIKE '62%' or CG_NUM LIKE '63%' or CG_NUM LIKE '65%'").sum(:ec_montant) 
#ap dbEc.select(:ec_montant).where("CG_NUM LIKE '61%'").sum(:ec_montant) 
#ap dbEc.select(:ec_montant).where( Sequel.|({:CG_NUM => ['61%']} )).sum(:ec_montant) 
#ap dbEc.select(:ec_montant).where( Sequel.like(:name, 'Acme%') ).sum(:ec_montant) 

# ne fonctionne pas car cherche des correspondances exactes et ne tiens donc pas compte des caracteres joker

#ap dbEc.select(:ec_montant).where( Sequel.|({:CG_NUM => ['61%', '62%', '63%', '65%']} )).filter{ (jm_date >= f>>m)  & (jm_date < f>> (m +1 )) }.sum(:ec_montant)


# Require the WIN32OLE library
require 'win32ole'
# Create an instance of the Excel application object
xl = WIN32OLE.new('Excel.Application')
# Make Excel visible
xl.Visible = 0
# Add a new Workbook object
wb = xl.Workbooks.Add
# Get the first Worksheet
ws = wb.Worksheets(1)
# Set the name of the worksheet tab
ws.Name = 'Result'




cpta = Compta.new

raise ArgumentError, "Date in database have a different format from queries", caller unless cpta.test_date? 

MONTH = options[:date].split("/")[1].to_i
YEAR = options[:date].split("/")[2].to_i
DAY = options[:date].split("/")[0].to_i

f = Date.new(YEAR, MONTH, DAY)
NBOFMONTH = options[:length].to_i

puts "date de comparatif début : #{DAY}/#{MONTH}/#{YEAR}"

puts "ALERTE BUG sur le taux de marge dernière ligne"


#http://stackoverflow.com/questions/14178439/how-to-tokenize-a-simple-mixed-string-into-either-ints-or-symbols
def tokenize(str)
  str.split(/(\d+%?)/).map! { |t| t[/\d/] ? t.to_s : t.strip.to_sym }
end


# test fichier configuration
require 'yaml'
sCfgFile = 'PandL_1.yml'
sheets = YAML.load_file(sCfgFile)

sheets.each do |i,sheet|
	puts "création du tableau #{i}" if options[:verbose]
	ws.Name = i
	#puts sheet.inspect
	iRowBegin = 1
	iColumnBegin = 1

	
	iRow = 0 # index de ligne courante
	iCol = 0 # index de colonne courante

	iLine = 0	# conserve le nombre de ligne du calcul en cours
	aLine = [];

	# PASS 1 (total and marge total must be treated in pass 2)		

	# display header
	ws.Cells(iRowBegin, iColumnBegin).Value = "Header"
	iCol=iColumnBegin
	NBOFMONTH.times do |m|
		ws.Columns( iCol+iColumnBegin ).NumberFormat = sFormatComptabilite
		ws.Cells(iRowBegin, iCol+iColumnBegin).NumberFormat = "mmm"
		ws.Cells(iRowBegin, iCol+iColumnBegin).Value = (cpta.addCivilMonth(f, m)).to_s	
		iCol+=1		
		if ( q = (m+1).modulo(3) ) === 0
			#trimestre
			ws.Columns( iCol+iColumnBegin ).NumberFormat = sFormatComptabilite
			ws.Cells(iRowBegin, iCol+iColumnBegin).Value = "T #{(m+1)/3}"
			ws.Range("#{(iCol+iColumnBegin-3).alph}:#{(iCol+iColumnBegin-1).alph}").Group
			iCol+=1
		end
		if ( q = (m+1).modulo(6) ) === 0
			#semestre
			ws.Columns( iCol+iColumnBegin ).NumberFormat = sFormatComptabilite
			ws.Cells(iRowBegin, iCol+iColumnBegin).Value = "S #{(m+1)/6}"
			iCol+=1
		end
		if ( q = (m+1).modulo(12) ) === 0
			ws.Columns( iCol+iColumnBegin ).NumberFormat = sFormatComptabilite
			ws.Cells(iRowBegin, iCol+iColumnBegin).Value = "A #{(m+1)/12}"
			iCol+=1			
		end
	end

	# Pour avoir trois niveaux de regroupement il faut faire la manoeuvre 3 fois

	# traiter les regroupements semestre
	#ws.Range("#{(iCol-3).alph}:#{(iCol-1).alph}").Group
	# traiter les regroupements annuels
	#ws.Range("#{(iCol-3).alph}:#{(iCol-1).alph}").Group

	iRow=iRowBegin+1

	puts "premiere passe" if options[:verbose]
	sheet.each do |v|
		#puts v.fetch(v.keys[0].to_i)[0]  identique à puts v.values[0][0]
		index =  v.keys[0].to_i
		libelle = v.values[0][0]
		compte = v.values[0][1].to_s
		type = v.values[0][2].to_s

		puts "traitement de la ligne #{index}" if options[:verbose]

		# display "libelle" at the position set in param
		iRow = (index > iRow ? index : iRow)
		iCol=iColumnBegin
		ws.Cells(iRow, iColumnBegin).Value = "#{libelle}"
		iCol+=1

		raise ArgumentError, "index #{index} is definited more than once in configuration file : #{sCfgFile}", caller unless aLine[index].nil? 

		if (type==="compte")				
			# prevent not signing if only one count
			compte = "+#{compte}" unless (compte[0] == "+" || compte[0] == "-")
			k = tokenize(compte)
			tot= []
			NBOFMONTH.times do |m|
				tot[m] = 0
				tott = 0
				# on renverse pour traiter les additions+soustractions en fin
				k.reverse.each do |tok|
					#ap "tok #{tok}"
					#ap tott
					if tok.is_a? Symbol
						if tok.to_s =="+"
							tot[m] += tott
						elsif tok.to_s =="-"
							tot[m] -= tott
						end
					else
						tott = cpta.c(tok, cpta.addCivilMonth(f, m),nil,true,nil,false)
					end
				end
			end
			
			NBOFMONTH.times do |m|
				ws.Cells(iRow, iCol).Value = tot[m]
				iCol+=1	
				iCol = cpta.do_total(ws, m, iRow, iCol)			
			end
			aLine[index] = iRow # les lignes vides peuvent être totalisées
			iRow+=1
		elsif (type === "vide")
			aLine[index] = iRow
			iRow+=1
		elsif (type==="mouvement")

			ws.Cells(iRow, iColumnBegin).Value = "Solde début de mois #{libelle}"
			ws.Cells(iRow+1, iColumnBegin).Value = "#{libelle} Débits"
			ws.Cells(iRow+2, iColumnBegin).Value = " #{libelle} Crédits"
			ws.Cells(iRow+3, iColumnBegin).Value = "Solde fin de mois #{libelle}"

			_saved_cur_column = iCol

			NBOFMONTH.times do |m|
				if m == 0 
					formula = 0.00
					puts "Attention pour le moment on utilise un hack mise à 0.00 en SAN il faudra calculer les an l'annee suivante voir avec les champs ecriture "
				else
					formula = "=SOMME(#{(_saved_cur_column-1).alph}#{iRow+3}"
				end
				ws.Cells(iRow, iCol).Formula = formula
				ws.Cells(iRow+1, iCol).Value = cpta.c(compte, cpta.addCivilMonth(f, m),nil,true,'D',false)
				ws.Cells(iRow+2, iCol).Value = cpta.c(compte, cpta.addCivilMonth(f, m),nil,true,'C',false)
				ws.Cells(iRow+3, iCol).Value = "=SOMME(#{(iCol).alph}#{iRow}+#{(iCol).alph}#{iRow+1}+#{(iCol).alph}#{iRow+2}"
				iCol+=1	
				_saved_cur_column = iCol
				if ( q = (m+1).modulo(3) ) === 0 #trimestre
					#ws.Columns('D:I').Interior.Color = 6
					iCol+=1
				end
				if ( q = (m+1).modulo(6) ) === 0 #semestre
					iCol+=1
				end
				if ( q = (m+1).modulo(12) ) === 0 #année
					iCol+=1
				end		

			end
			aLine[index] = iRow+3
			iRow+=4
		elsif (type==="liste")	
			#ws.Cells(iRow, iColumnBegin).Value = "#{libelle}"
			iSaveRow = iRow # keep memories of first line
			iRow+=1 # pass first tot line
			list_c = cpta.l(compte, f , NBOFMONTH)
			list_c.each do |c|
				iCol=iColumnBegin	# remise de l'index colonne en debut de colonne
				ws.Cells(iRow, iCol).Value = sprintf("%d-%s",c[0], cpta.l_compte(c[0]))	
				iCol+=1
				NBOFMONTH.times do |m|
					ws.Cells(iRow, iCol).Value = c[1][m]
					iCol+=1	
					iCol = cpta.do_total(ws,m,iRow,iCol)	
				end
				iRow+=1
			end
			# do first line total
		    aLine[index] = iSaveRow
			iCol=iColumnBegin+1	# remise de l'index colonne en debut de colonne +1 (libelle deja ecrit)
			NBOFMONTH.times do |m|
				if list_c.empty?
					ws.Cells(iSaveRow, iCol).Value = 0.00
				else
					ws.Cells(iSaveRow, iCol).Formula = "=SOMME(#{(iCol).alph}#{iSaveRow+1}:#{(iCol).alph}#{iRow-1})"
				end
				iCol+=1	
				iCol = cpta.do_total(ws,m,iSaveRow,iCol)			
			end	
		elsif (type==="ratio")
			aLine[index] = iRow

		elsif (type==="total")
			aLine[index] = iRow
			iRow+=1
		else
			raise RuntimeError, "Type #{type} specified in yaml file is not treated by this application", caller
		end
	end
	puts "seconde passe" if options[:verbose]

	iRow = iRowBegin
	
	sheet.each do |v|

		index 	= v.keys[0]
		libelle = v.values[0][0]
		compte 	= v.values[0][1].to_s
		type 	= v.values[0][2].to_s
		
		puts "traitement de la ligne #{index}" if options[:verbose]

		iCol = iColumnBegin
		if (type==="total")
			k = tokenize(compte)
			formula = String.new
			k.each do |tok|
				if tok.is_a? Symbol
					if tok.to_s =="+"
						formula += "+"
					elsif tok.to_s =="-"
						formula += "-"
					end
				else
					raise RuntimeError, "Index #{tok.to_i} not exists in ...", caller if aLine[tok.to_i].nil? 
					formula += "__col__#{aLine[tok.to_i]}"
				end
			end			
			iRow=aLine[index]
			iCol+=1 # libelle already putted
			NBOFMONTH.times do |m|
				ws.Cells(iRow, iCol).Formula = "=#{formula}".gsub(/__col__/, (iCol).alph ) 
				iCol+=1	
				iCol = cpta.do_total(ws,m,iRow,iCol)		
			end
		elsif (type==="ratio")
			num = compte.split("/")[0]
			den = compte.split("/")[1]
			formula1 = cpta.formula_total(cpta.tokenize2(num),aLine)
			formula2 = cpta.formula_total(cpta.tokenize2(den),aLine)

# -- fonctionne mais complique tout
=begin
			tok = "7%"
			m = 12
			a = "cpta.c(tok, f, f>>m, true,nil,true)"
			puts a
			b = eval(a)
			puts b

			a =  formula2.gsub(/C(\d+%?)/,'cpta.c(\1, f, f>>m, true,nil,true)')
			puts a
			b = eval(a)
			puts b
=end
			iRow=aLine[index]
			iCol+=1 # libelle already putted
			NBOFMONTH.times do |m|

				formula = "si(#{formula2.gsub(/__col__/, (iCol).alph )}=0;0;"
				formula += "(#{formula1.gsub(/__col__/, (iCol).alph )})"
				formula += "/(#{formula2.gsub(/__col__/, (iCol).alph )}))"


				ws.Cells(iRow,iCol).NumberFormat = sFormatPourcentage
				ws.Cells(iRow, iCol).Formula = "=#{formula}"

				iCol+=1	
				iCol = cpta.do_total(ws, m, iRow, iCol,true)			
			end
			aLine[index] = iRow
			iRow+=1
		end

	end
end

puts "L'application a executé #{cpta.totSQL} requêtes SQL pour fournir le resultat"



#saving as pre-2007 format
excel97_2003_format = -4143 
pwd =  Dir.pwd.gsub('/','\\') << '\\'
wb.SaveAs("#{pwd}#TEST_#{YEAR}#{MONTH}#{DAY}.xls", excel97_2003_format)
# Save the workbook
#wb.SaveAs("#{pwd}#{sFileName}.xls")
# Close the workbook
wb.Close
# Quit Excel
xl.Quit
WIN32OLE.ole_free(xl)

exit

def object_count
  count = Hash.new(0)
  ObjectSpace.each_object{|o| count[o.class] += 1}
  count
end

require 'pp'


pp object_count
GC.end
