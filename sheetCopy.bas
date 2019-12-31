
private sub cmdbt_click()
label1.caption=activeworkbook.name
end sub

private sub cmdbt_click()

dim st as worksheet

for each st activewindow.selectedsheets
st.copy before:=workbooks(label1.caption).sheets(activeworkbook.sheets.count)
next


end sub
