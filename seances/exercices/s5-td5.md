<p style="text-align:left;">
    [Retour au sommaire](../../README.md)
    <span style="float:right;">
        [Séance 4 - VBA: Structures itératives et communication avec Excel](s4-vba-2.md)
    </span>
</p>

<div style="text-align:center;">
# Correction du TD 5 - Tableaux et Enregistrements
</div>

---

## Exercice 1

Ecrire un programme permettant:

* de déclarer un tableau de taille au plus **TMAX=100**;
* de saisir les éléments de ce tableau;
* d'afficher les éléments de ce tableau.

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

#### Correction

<details>

```vb
Const tmax = 100

Sub Ex_1()
Dim n As Integer, i As Integer
Dim t(1 To tmax) As Double

n = InputBox("Combien de valeurs voulez-vous saisir ?")

If n <= 0 Or n > tmax Then
    MsgBox ("Erreur de saisie")
Else
    For i = 1 To n
        t(i) = InputBox("Entrez la " & i & " ème valeur du tableau")
    Next
    
    For i = 1 To n
        MsgBox ("t[" & i & "] = " & t(i))
    Next
End If
End Sub
```

</details>
</div>

---

## Exercice 2

Ecrire un programme permettant:

* de saisir un tbleau de cinq entier;
* d'afficher les éléments de ce tableau en commençant par le dernier qui a été saisi.

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

#### Correction

Si l'on souhaite utiliser des boucles **For**:

<details>

```vb
Sub Ex_2()
    Dim t(1 To 5) As Integer
    Dim i As Integer

    t(1) = InputBox("Entrez le premier entier")

    For i = 2 To 5
        t(i) = InputBox("Entrez le " & i & " ème entier")
    Next i

    For i = 5 To 1 Step -1
        MsgBox ("t[" & i & "]= " & t(i))
    Next i
End Sub
```

</details>

Si l'on souhaite utiliser uniquement des boucles **While**:

<details>

```vb
Sub Ex_2()
    Dim t(1 To 5) As Integer
    Dim i As Integer

    t(1) = InputBox("Entrez le premier entier")

    i = 2
    While i <= 5
        t(i) = InputBox("Entrez le " & i & " ème entier")
        i = i + 1
    Wend

    i = 5
    While i >= 1
        MsgBox ("t[" & i & "]= " & t(i))
        i = i - 1
    Wend
End Sub
```

</details>

Si l'on souhaite utiliser uniquement des boucles **Do While**:

<details>

```vb
Sub Ex_2()
    Dim t(1 To 5) As Integer
    Dim i As Integer

    i = 1
    Do While i <= 5
        t(i) = InputBox("Entrez le " & i & " ème entier")
        i = i + 1
    Loop

    i = 5
    Do While i >= 1
        MsgBox ("t[" & i & "]= " & t(i))
        i = i - 1
    Loop
End Sub
```

</details>

</div>

---

## Exercice 3

Ecrire un programme permettant:

* de saisir un tableau de cinq entier;
* d'échanger les valeurs de la première et de la dernière case si la première valeur est plus grande que la dernière.

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

#### Correction

<details>

```vb
Sub Ex_3()
    Dim t(1 To 5) As Integer
    Dim i As Integer
    Dim tmp As Integer
    t(1) = InputBox("Entrez le premier entier")

    For i = 2 To 5
    t(i) = InputBox("Entrez le " & i & " ème entier")
    Next

    ' On échange les valeurs avec une variable temporaire'
    If t(1) > t(5) Then
        tmp = t(1)
        t(1) = t(5)
        t(5) = tmp
    End If

    ' On vérifie que les valeurs ont été échangées'
    For i = 1 To 5
        MsgBox ("t[" & i & "]= " & t(i))
    Next
End Sub
```

</details>

</div>

---

## Exercice 4

Ecrire un programme permettant:

* de déclarer un tableau de taille au plus **TMAX=100**;
* de déterminer le nombre d'éléments saisis sur la première ligne de la *Feuille1* du classeur *Excel*;
* de stocker les éléments situés sur la première ligne de la *Feuille1* du classeur *Excel* dans ce tableau;
* d'afficher les éléments de ce tableau sur la première colonne de la *Feuille2* du classeur *Excel*.
d'échanger les valeurs du tableau de la *Feuille1* si la première valeur est plus grande que la dernière. On affichera ce tableau *trié* sur la deuxième ligne de la *Feuille1*.

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

#### Correction

<details>

```vb
Sub Ex_4()
    Dim t(1 To tmax) As Double
    Dim i As Integer
    Dim nc As Integer
    Dim tmp As Double

    i = 1

    ' On détermine le nombre d'éléments saisis'
    While Feuil1.Cells(1, i) <> "" And i <= tmax
        i = i + 1
    Wend

    nc = i - 1

    ' On stocke les éléments dans le tableau'
    For i = 1 To nc
        t(i) = Feuil1.Cells(1, i)
    Next

    ' On affiche les éléments du tableau'
    For i = 1 To nc
        Feuil2.Cells(i, 1) = t(i)
    Next

    ' On échange les valeurs avec une variable temporaire'
    If t(1) > t(nc) Then
        tmp = t(1)
        t(1) = t(nc)
        t(nc) = tmp
    End If

    'On affiche le tableau trié'
    For i = 1 To nc
        Feuil1.Cells(2, i) = t(i)
    Next
End Sub
```

</details>

</div>