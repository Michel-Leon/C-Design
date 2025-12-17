Attribute VB_Name = "M�dulo1"

public sub Grupo_Dibujo()
    dim shp as shape
    dim groupname as string
    groupname = "Grupo Dibujo"

    on error resume next
    set shp = ActiveSheet.Shapes(groupname)
    on error goto 0

    if not shp is nothing then
        shp.visible = not shp.visible
        ActiveSheet.Shapes("Grupo plano corte").visible = false
        ActiveSheet.shapes("Grupo plano dobles").visible = false
        ActiveSheet.Shapes("Grupo plano Ensamble").visible = false
        ActiveSheet.shapes("Grupo plano Soldadura").visible = false
    end if
end sub

public sub Grupo_Plano_Corte()
    dim shp as shape
    dim groupname as string
    groupname = "Grupo plano corte"

    on error resume next
    set shp = ActiveSheet.Shapes(groupname)
    on error goto 0

    if not shp is nothing then
        shp.visible = not shp.visible
        ActiveSheet.Shapes("Grupo Dibujo").visible = false
        ActiveSheet.shapes("Grupo plano dobles").visible = false
        ActiveSheet.Shapes("Grupo plano Ensamble").visible = false
        ActiveSheet.shapes("Grupo plano Soldadura").visible = false
    end if
end sub
public sub Grupo_Plano_Dobles()
    dim shp as shape
    dim groupname as string
    groupname = "Grupo plano dobles"

    on error resume next
    set shp = ActiveSheet.Shapes(groupname)
    on error goto 0

    if not shp is nothing then
        shp.visible = not shp.visible
        ActiveSheet.Shapes("Grupo Dibujo").visible = false
        ActiveSheet.shapes("Grupo plano corte").visible = false
        ActiveSheet.Shapes("Grupo plano Ensamble").visible = false
        ActiveSheet.shapes("Grupo plano Soldadura").visible = false
    end if
end sub
public sub Grupo_Plano_Ensamble()
    dim shp as shape
    dim groupname as string
    groupname = "Grupo plano Ensamble"

    on error resume next
    set shp = ActiveSheet.Shapes(groupname)
    on error goto 0

    if not shp is nothing then
        shp.visible = not shp.visible
        ActiveSheet.Shapes("Grupo Dibujo").visible = false
        ActiveSheet.shapes("Grupo plano corte").visible = false
        ActiveSheet.Shapes("Grupo plano dobles").visible = false
        ActiveSheet.shapes("Grupo plano Soldadura").visible = false
    end if
end sub
public sub Grupo_Plano_Soldadura()
    dim shp as shape
    dim groupname as string
    groupname = "Grupo plano Soldadura"

    on error resume next
    set shp = ActiveSheet.Shapes(groupname)
    on error goto 0

    if not shp is nothing then
        shp.visible = not shp.visible
        ActiveSheet.Shapes("Grupo Dibujo").visible = false
        ActiveSheet.shapes("Grupo plano corte").visible = false
        ActiveSheet.Shapes("Grupo plano dobles").visible = false
        ActiveSheet.shapes("Grupo plano Ensamble").visible = false
    end if
end sub