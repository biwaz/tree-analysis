function create_tree(dest, self, parent)
	dim objDict, temp(), i, n
	set objDict = createobject("Scripting.Dictionary")
	redim temp(ubound(dest))
	for i = 0 to ubound(dest)
		objDict(dest(i)(self)) = i
	next
	for i = 0 to ubound(dest)
		if not objDict.exists(dest(i)(parent)) then
			n = -1
		else
			if dest(i)(self) = dest(i)(parent) then n = -1 else n = objDict(dest(i)(parent))
		end if
		dest(i)(parent) = n
		temp(i) = cstr(n)
	next
	for i = 0 to ubound(dest)
		n = dest(i)(parent)
		if 0 <= n then temp(n) = temp(n) & "," & i
	next
	for i = 0 to ubound(dest)
		dest(i)(parent) = split(temp(i), ",")
	next
	set create_tree = objDict
end function
