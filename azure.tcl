ttk::style theme create azure -parent clam -settings {
    # Define default colors
    set bgColor #f0f0f0
    set fgColor #333333
    set btnColor #0078D7
    set btnActiveColor #005FA7
    set btnTextColor white

    # Button style
    ttk::style configure TButton -background $btnColor -foreground $btnTextColor -font {Helvetica 10} -padding {15 15}
    

    # Frame style
    ttk::style configure TFrame -background $bgColor

    # LabelFrame style
    ttk::style configure TLabelframe -background $bgColor

    # Label style
    ttk::style configure TLabel -background $bgColor -foreground $fgColor -font {Helvetica 10}

    # Entry style
    ttk::style configure TEntry -fieldbackground white -foreground $fgColor -font {Helvetica 10}

    # Treeview style
    ttk::style configure TTreeview -background white -fieldbackground white -foreground $fgColor -font {Helvetica 10}
}

ttk::style theme use azure
