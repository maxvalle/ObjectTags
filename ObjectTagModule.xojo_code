#tag Module
Protected Module ObjectTagModule
	#tag Method, Flags = &h0
		Function ObjectTag(extends o as Object) As Variant
		  //the main purpose of the pragma is to prevent a context switch.
		  #pragma disableBackgroundTasks
		  
		  Tags = RemoveDeadRefs(Tags)
		  
		  dim tagValue as Variant
		  for each p as Pair in Tags
		    dim w as WeakRef = p.Left
		    if w.Value = o then
		      tagValue = p.Right
		      exit
		    end if
		  next
		  
		  return tagValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ObjectTag(extends o as Object, assigns value as Variant)
		  //the main purpose of the pragma is to prevent a context switch.
		  #pragma disableBackgroundTasks
		  
		  Tags = RemoveDeadRefs(Tags)
		  
		  dim index as Integer = -1
		  for i as Integer = 0 to UBound(Tags)
		    dim w as WeakRef = Tags(i).Left
		    if w.Value = o then
		      index = i
		      exit
		    end if
		  next
		  
		  if index > -1 then
		    Tags(index) = new WeakRef(o) : value
		  else
		    Tags.Append new WeakRef(o) : value
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RemoveDeadRefs(tagList() as Pair) As Pair()
		  //the main purpose of the pragma is to prevent a context switch.
		  #pragma disableBackgroundTasks
		  
		  dim newList() as Pair
		  
		  for each p as Pair in tagList
		    dim w as WeakRef = p.Left
		    if w.Value <> nil then
		      newList.Append p
		    end if
		  next
		  
		  return newList
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private tags() As Pair
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
