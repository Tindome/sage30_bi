=begin
	
	Some link
		https://groups.google.com/forum/#!topic/sketchupruby/cBEOY96g7I0 => help with excel ole properties


	VERSION
	1.00	+ (améliorations)
			- (correction de bug)
	1.01    - dans la version précédente les dates étaient exprimées au format YYYY/MM/DD, désormais suite à l'utilisation de create_database, les dates sont exprimées
			au format YYYY-MM-DD, on change donc les formules de requetes strftime('%Y/%m/%d' remplacé par strftime('%Y-%m-%d'
			+ mise en place d'un test sur les dates pour vérifier ce genre de bétises

	TODO


	EXPLANATION
=end 

class Compta

	
	DB= Sequel.connect('sqlite://db/sage.db', :logger => Logger.new('log/db.log'))

    def initialize(name="dummy")
        @name = name 
        @iTotSQLRequest =0
    end  

    # return FALSE if date is YYYY/MM/DD TRUE if date is YYYY-MM-DD
    def test_date?
    	DB[:F_ECRITUREC].select(:JM_DATE).first[:JM_DATE].to_s.match(/-/).to_s === "-"
    end

	def xl_total(ws,row,column,start_r,start_c,end_r=nil,end_c=nil)
		#end_r = (end_r or start_r)
		end_r ||= start_r
		#end_c = (end_c or start_c)
		end_c ||= start_c
		ws.Cells(row, column).Formula = "=SOMME(#{(start_r).alph}#{start_c}:#{(end_r).alph}#{end_c})"
	end

	# ws 			: worksheet
	# m 			: month (1..n)
	# iRow, iCol 	: reference of current row and current col

	def do_total(ws,m,iRow,iCol,bDoAverage=false,bDbg=false)

		if (bDbg)
			p "mois #{m}"
			p "est un trimestre" if ( q = (m+1).modulo(3) ) === 0
			p "est un semestre" if ( q = (m+1).modulo(6) ) === 0
			p "est une année" if ( q = (m+1).modulo(12) ) === 0
			p "Row #{iRow}, Col #{iCol}"
		end
		
		if ( q = (m+1).modulo(3) ) === 0 #trimestre
			if (false === bDoAverage)
				self.xl_total(ws,iRow, iCol,iCol-3,iRow,iCol-1,iRow)
			else
				ws.Cells(iRow, iCol).Formula = "=si(3-NB.SI(#{(iCol-3).alph}#{iRow}:#{(iCol-1).alph}#{iRow};0)=0;0;(#{(iCol-3).alph}#{iRow}+#{(iCol-2).alph}#{iRow}+#{(iCol-1).alph}#{iRow}) / (3-NB.SI(#{(iCol-3).alph}#{iRow}:#{(iCol-1).alph}#{iRow};0)))"
			#ws.Columns('D:I').Interior.Color = 6
			end
			ws.Cells(iRow,iCol).NumberFormat = ws.Cells(iRow,iCol-1).NumberFormat unless ws.Cells(iRow,iCol-1).NumberFormat.length == 0
			iCol+=1
		end
		if ( q = (m+1).modulo(6) ) === 0 #semestre
			if (false === bDoAverage)
				ws.Cells(iRow, iCol).Formula = "=#{(iCol-5).alph}#{iRow}+#{(iCol-1).alph}#{iRow}"
			else
				ws.Cells(iRow, iCol).Formula = "=si(#{(iCol-5).alph}#{iRow}+#{(iCol-1).alph}#{iRow}=0;0;si(#{(iCol-5).alph}#{iRow}=0;#{(iCol-1).alph}#{iRow};si(#{(iCol-1).alph}#{iRow}=0;#{(iCol-5).alph}#{iRow};(#{(iCol-5).alph}#{iRow}+#{(iCol-1).alph}#{iRow})/2)))"
			end
			#precaution
			ws.Cells(iRow,iCol).NumberFormat = ws.Cells(iRow,iCol-1).NumberFormat unless ws.Cells(iRow,iCol-1).NumberFormat.length == 0
			iCol+=1
		end
		if ( q = (m+1).modulo(12) ) === 0 #année
			if (false === bDoAverage)
				ws.Cells(iRow, iCol).Formula = "=#{(iCol-10).alph}#{iRow}+#{(iCol-1).alph}#{iRow}"
			else
				ws.Cells(iRow, iCol).Formula = "=si(#{(iCol-10).alph}#{iRow}+#{(iCol-1).alph}#{iRow}=0;0;si(#{(iCol-10).alph}#{iRow}=0;#{(iCol-1).alph}#{iRow};si(#{(iCol-1).alph}#{iRow}=0;#{(iCol-10).alph}#{iRow};(#{(iCol-10).alph}#{iRow}+#{(iCol-1).alph}#{iRow})/2)))"
			end
			#ap ws.Cells(iRow,iCol-1).NumberFormat
			#precaution
			ws.Cells(iRow,iCol).NumberFormat = ws.Cells(iRow,iCol-1).NumberFormat unless ws.Cells(iRow,iCol-1).NumberFormat.length == 0
			iCol+=1
		end	
		return iCol
	end

	def cs1(num_compte)
		return c(num_compte,dtDd,dtDb>>1,solde=true,sens=nil)
	end

	def c(num_compte,dtDd,dtDf=nil,solde=true,sens=nil,dbg=false)
		dbEc = DB[:F_ECRITUREC]
		dtDf ||= addCivilMonth(dtDd,1)

		if num_compte.class == Array
			sCompte = num_compte.inject(""){ |s,c| s+" CG_NUM LIKE '#{c}' OR "}
			#enlever les 3 derniers char ("or ")
			sCompte.slice!(-3..-1)
		else
			sCompte ="CG_NUM LIKE '#{num_compte}'"
		end

		if solde
			c = dbEc.select(:EC_MONTANT).where("(#{sCompte}) AND EC_SENS=1").filter{ (jm_date >= strftime('%Y-%m-%d',dtDd))  & (jm_date < strftime('%Y-%m-%d',dtDf)) }.sum(:EC_MONTANT)
			p dbEc.select(:EC_MONTANT).where("(#{sCompte}) AND EC_SENS=1").filter{ (jm_date >= strftime('%Y-%m-%d',dtDd))  & (jm_date < strftime('%Y-%m-%d',dtDf)) }.sql if dbg===true
			@iTotSQLRequest+=1
			c = 0.00 if c.nil?
			d = dbEc.select(:EC_MONTANT).where("(#{sCompte}) AND EC_SENS=0").filter{ (jm_date >= strftime('%Y-%m-%d',dtDd))  & (jm_date < strftime('%Y-%m-%d',dtDf)) }.sum(:EC_MONTANT)
			p dbEc.select(:EC_MONTANT).where("(#{sCompte}) AND EC_SENS=0").filter{ (jm_date >= strftime('%Y-%m-%d',dtDd))  & (jm_date < strftime('%Y-%m-%d',dtDf)) }.sql if dbg===true
			@iTotSQLRequest+=1
			d = 0.00 if d.nil?
			if sens.nil? 
				#m = dbEc.select(:ec_montant).where("#{sCompte}").filter{ (jm_date >= dtDd)  & (jm_date < dtDf) }.sum(:ec_montant)
				#m.nil? ? m = 0.00 : nil
				#return m.to_f
				d.to_f - c.to_f
			else
				if sens.upcase() === "C"
					#m = dbEc.select(:ec_montant).where("(#{sCompte}) AND EC_SENS=1").filter{ (jm_date >= dtDd)  & (jm_date < dtDf) }.sum(:ec_montant)
					#m.nil? ? m = 0.00 : nil
					#return m.to_f
					- c.to_f
				else
					#m = dbEc.select(:ec_montant).where("(#{sCompte}) AND EC_SENS=0").filter{ (jm_date >= dtDd)  & (jm_date < dtDf) }.sum(:ec_montant)
					#m.nil? ? m = 0.00 : nil
					#return m.to_f
					d.to_f
				end
			end
		else
			# attention bug à la con, si la requete SQL n'est pas la derniere ligne on ne retourne pas le nombre de résultats trouvés
			@iTotSQLRequest+=1
			p dbEc.select(:JM_DATE,:CG_NUM,:EC_INTITULE,:EC_MONTANT,:EC_SENS ).where("#{sCompte}").filter{ (jm_date >= strftime('%Y-%m-%d',dtDd))  & (jm_date < strftime('%Y-%m-%d',dtDf)) }.sql if dbg===true
			dbEc.select(:JM_DATE,:CG_NUM,:EC_INTITULE,:EC_MONTANT,:EC_SENS ).where("#{sCompte}").filter{ (jm_date >= strftime('%Y-%m-%d',dtDd))  & (jm_date < strftime('%Y-%m-%d',dtDf)) }
		end 
	end

	def l(compte,dtDd,iMonth=1,dbg=false)
		dtDf ||= self.addCivilMonth(dtDd,iMonth)

		if (compte[0] == "+" || compte[0] == "-")
			operation = compte[0] 
			compte = compte[1..-1] 
		else 
			operation = "+"
		end

		ds = self.c(compte, dtDd , dtDf, false, nil, dbg)
		list= Hash.new
		
		ds.all.each do |p|
			#d,c,i,m,s = p[:jm_date],p[:cg_num],p[:ec_intitule],p[:ec_montant],p[:ec_sens]
			d,c,i,m,s = p[:JM_DATE],p[:CG_NUM],p[:EC_INTITULE],p[:EC_MONTANT],p[:EC_SENS]
			list[p[:CG_NUM]].nil? ? list[p[:CG_NUM]] = Array.new(iMonth, 0.00) : nil
			p[:EC_SENS] ?  m=-m : true# si sens crediteur cad EC_SENS = 1 mettre le montant en negatif
			# save in 0 the sum of all months
			#operation == "+" ? list[c][0]+=m : list[c][0]-=m 
			operation == "+" ? list[c][self.date_diff_m(dtDd,d)]+=m : list[c][self.date_diff_m(dtDd,d)]-=m
		end
		
		return list
	end

	def l_compte(num_compte)
		@iTotSQLRequest+=1
		DB[:F_COMPTEG].select(:CG_INTITULE).where("CG_NUM LIKE '#{num_compte}'").first[:CG_INTITULE]
		
	end

	#add exactly one month ... because >> add one month but not changing the day ie 30/11 + 1 month = 30/12
	def addCivilMonth(dt,iMonth=1)
		( ( dt + 1 ) >> iMonth ) - 1
	end

	#http://stackoverflow.com/questions/14178439/how-to-tokenize-a-simple-mixed-string-into-either-ints-or-symbols
	def tokenize(str)
	  str.split(/(\d+%?)/).map! { |t| t[/\d/] ? t.to_s : t.strip.to_sym }
	end
	# this tokenize2 function give us the ability to extract CXXX part in tokenize C is a token and XXX another
	def tokenize2(str)
	  str.split(/([+|-|\/|\*|C\d]+%?)/).map! { |t| t[/\d/] ? t.to_s : t.strip.to_sym }
	end		

	##
	# src : http://stackoverflow.com/questions/18892657/get-no-of-months-years-between-two-dates-in-ruby 
	# STT just need number of months so modified it a bit
	# Calculates the difference in years and month between two dates
	# Returns number of months
	def date_diff_m(date1,date2)
	  (date2.year * 12 + date2.month) - (date1.year * 12 + date1.month)
	end

	def totSQL
		@iTotSQLRequest
	end

	def formula_total(k,aLine)
		formula = String.new
		partial = String.new
		k.each do |tok|
			next if tok.length == 0 
			if tok.is_a? Symbol
				if tok.to_s =="+"
					formula += "+"
				elsif tok.to_s =="-"
					formula += "-"
				else
					raise RuntimeError, "Symbol #{tok} not treated", caller
				end
			elsif tok.include?("C")
				raise RuntimeError, "Calculating account soldes in total isn t performed at this time in formula #{tok} in yaml file", caller
			else 
				raise RuntimeError, "Index #{tok.to_i} not exists in ...", caller if aLine[tok.to_i].nil? 
				formula += "__col__#{aLine[tok.to_i]}"
			end
		end	
		formula
	end

end