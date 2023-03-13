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

---

## Exercice 5

On considère la feuille de calcul (*Feuille3*) suivante:

![s5-exo5](screenshots/s5-exo5.png)

Ecrivez en VB un programme permettant de calculer la valeur totale du stock (simme des produits des prix par les quantités en stock) et de l'inscrire dans la cellule **F4**.

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

#### Correction

<details>

```vb
Sub Ex_5()

    Dim t(1 To 5, 1 To 2) As Double
    Dim i As Integer
    Dim somme As Double

    somme = 0

    For i = 1 To 5
        t(i, 1) = Feuil3.Cells(i + 1, 2)
        t(i, 2) = Feuil3.Cells(i + 1, 3)
    Next

    For i = 1 To 5
        somme = somme + t(i, 1) * t(i, 2)
    Next
    Feuil3.Cells(4, 6) = Format(somme, "# €")
End Sub

```

</details>

</div>

---

## Exercice 6

Recopiez le code ci-dessous puis expliquez ce qu'il réalise.

```vb
Const tmax = 100

Sub exo6()
    Dim t(1 To tmax) As Integer
    Dim i As Integer
    Dim nc As Integer
    Dim tmp As Integer
    Dim permut As Integer

    permut = 1
    i = 1

    While Feuil1.Cells(1, i) <> "" And i <= tmax
        i = i + 1
    Wend

    nc = i - 1

    For i = 1 To nc
        t(i) = Feuil1.Cells(1, i)
    Next

    While permut = 1
        permut = 0
        For i = 1 To nc - 1
            If t(i) > t(i + 1) Then
                permut = 1
                tmp = t(i)
                t(i) = t(i + 1)
                t(i + 1) = tmp
            End If
        Next
    Wend

    For i = 1 To nc
        Feuil1.Cells(2, i) = t(i)
    Next
End Sub

```

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

#### Correction

<details>

Le programme ci-dessus permet de trier un tableau de nombres entiers en ordre croissant.

</details>

</div>

---

## Exercice 7


<div style="border-left:solid #17a589 4px;padding-left:10px; ">

#### Correction

<details>

Créer une liste de chaînes de caractères comportant les éléments "Le ", "printemps", "arrive."

Ecrire ensuite une procédure affichant, à l'aide de la liste, la phrase: "Le printemps arrive".

```vb
Sub Ex_7()
    Dim l As Variant, phrase As String, i As Integer
    l = Array("Le", "printemps", "arrive")
    phrase = ""
    For i = 0 To 2
        phrase = phrase & Space(1) & l(i)
    Next i
    MsgBox phrase
End Sub
```

</details>

</div>
