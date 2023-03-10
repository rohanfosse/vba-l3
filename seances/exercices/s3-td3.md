<p style="text-align:left;">
    [Retour au sommaire](../../README.md)
</p>

<div style="text-align:center;">
# Correction du TD 3
</div>

## Exercice 1

---

#### Question 1

Ecrire une fonction nommée `taux`renvoyant le taux du livret A à 3,5%.
(la fonction sera testée par un appel de fonction dans une procédure dédiée à cela, ou directement sur une feuille de calcul)

##### Correction

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

Pour écrire cette fonction, il suffit de déclarer une fonction `taux` qui renvoie un
nombre à virgule flottante (un Single en VBA), et de renvoyer la valeur `0,035`.

Pour rappel, si l'on souhaite renvoyer une valeur dans une fonction en VBA, il faut
utiliser une variable du **même nom** que la fonction, et lui affecter la valeur à renvoyer.

La procédure sera simplement composée d'un appel de fonction, et d'un affichage de la
valeur renvoyée par la fonction.

Voici une solution possible:


<details>

```php
'Question 1'
Function taux() As Single
    taux = 0.035
End Function

Sub appel_taux()
    MsgBox ("le taux est de " & taux())
End Sub
```

</details>
</div>


---

#### Question 2

Ecrire une fonction nommée `tauxAn` prenant en paramètre d'entrée un nombre entier réprésentant une année, et renvoyant le taux du livret A e l'année, sachant que ce taux est de;

- 0,75% de 2017 à 2019 inclus;
- 0,5% en 2020 et en 2021;
- 2% en 2022;
- 3% en 2023.

Un structure conditionnelle `If ... Then ...`sera mise en oeuvre dans cette question.

##### Correction

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

Dans cette question, il est demandé de créer une fonction prenant en entrée un nombre entier (`l'année`), et renvoyant un nombre à virgule flottante (`le taux`).

Pour rappel, les paramètres d'une fonction sont déclarés entre parenthèses après le nom de la fonction, et sont séparés par une virgule. Les paramètres sont des variables locales à la fonction, et sont utilisés comme des variables classiques.

Pour plus d'informations, lire la section du cours sur les fonctions [ici](README.md#-les-fonctions).

Il est de plus demandé d'utiliser une structure conditionnelle `If ... Then ...` pour déterminer le taux en fonction de l'année.

Pour plus d'informations, lire la section du cours sur les structures conditionnelles [ici](README.md#-les-conditions).


<details>

```php
'Question 2'
Function tauxAn(a As Integer) As Single
    If a >= 2017 And a <= 2019 Then
        tauxAn = 0.0075
    ElseIf a >= 2020 And a <= 2021 Then
        tauxAn = 0.005
    ElseIf a = 2022 Then
        tauxAn = 0.02
    ElseIf a = 2023 Then
        tauxAn = 0.03
    End If
    
End Function

Sub appel_tauxAn()    'Procédure d appel de la fonction'
    MsgBox ("La remise finale est égale à " & tauxAn(2022))
End Sub
```

</details>

</div>

---

#### Question 3

Reprendre la question 2 avec la structure `Select Case`

##### Correction

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

Pour rappel, la structure `Select Case` permet de tester une variable, et d'exécuter un bloc de code en fonction de la valeur de cette variable.

Pour plus d'informations, lire la section du cours sur les `Select Case` [ici](README.md#-les-conditions).

<details>

```php
'Question 3'
Function tauxAnSelect(a As Integer) As Single
    Select Case a
        Case 2017 To 2019
            tauxAnSelect = 0.0075
        Case 2020, 2021
            tauxAnSelect = 0.005
        Case 2022
            tauxAnSelect = 0.02
        Case 2023
            tauxAnSelect = 0.03
    End Select
End Function

Sub appel_tauxAnSelect()    'Procédure d appel de la fonction'
    MsgBox ("La remise finale est égale à " & tauxAnSelect(2022))
End Sub
```

</details>

</div>

---

#### Question 4

Reprendre la question 2, en ajoutant le message texte `Inconnu` renvoyé par la fonction si l'année est inférieure à 2017 ou supérieure à 2023.

##### Correction

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

Dans cette question, il est demandé  à la fonction de renvoyer deux types de valeurs différentes, un nombre à virgule flottante, et une chaîne de caractères en fonction de la valeur de l'année.

En VBA, il existe un type de variable `Variant`, qui permet de stocker n'importe quel type de valeur. Il est donc possible de renvoyer une valeur de type `Variant` dans une fonction.

Voici une solution possible:

<details>

```php
'Question 4'

Function tauxAnInconnu(a As Integer) As Variant
    If a >= 2017 And a <= 2019 Then
        tauxAnInconnu = 0.0075
    ElseIf a >= 2020 And a <= 2021 Then
        tauxAnInconnu = 0.005
    ElseIf a = 2022 Then
        tauxAnInconnu = 0.02
    ElseIf a = 2023 Then
        tauxAnInconnu = 0.03
    Else
        tauxAnInconnu = "Inconnu"
    End If
    
End Function

Sub appel_tauxAnInconnu()    'Procédure d appel de la fonction'
    MsgBox ("La remise finale est égale à " & tauxAnInconnu(2022))
End Sub
```

</details>

</div>

---

#### Question 5

Reprendre l'exercice en codant cette fois une procédure dans laquelle une fenêtre contextuelle demande à l'utilisateur e saisir une année comprise entre 2017 et 2023, puis affiche en sortie d'écran le taux de l'année saisie. Si l'année saisie n'est pas dans l'intervalle demandé, le programme prend fin.

##### Correction

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

Dans cette question, il est demandé de créer une procédure, et non une fonction.

Pour rappel, les procédures sont des fonctions sans valeur de retour. Elles sont déclarées avec le mot clé `Sub` au lieu de `Function`.

Pour demander à l'utilisateur de saisir une valeur, il faut utiliser la fonction `InputBox`. Cette fonction renvoie la valeur saisie par l'utilisateur sous forme de chaîne de caractères.

Pour quitter une procédure, il faut utiliser la commande `Exit Sub`.

Une solution possible est la suivante:

<details>

```php
'Question 5'
Sub question_5()
    Dim a As Integer
    Dim taux As Double
    a = InputBox("Veuillez saisir une année comprise entre 2017 et 2023")
    If a < 2017 Or a > 2023 Then
        Exit Sub
    End If
    If a >= 2017 And a <= 2019 Then
        taux = 0.0075
    ElseIf a = 2020 Or a = 2021 Then
        taux = 0.005
    ElseIf a = 2022 Then
        taux = 0.02
    ElseIf a = 2023 Then
        taux = 0.03
    End If
    MsgBox "Le taux de l'année " & a & " est de: " & taux  'Il s agit ici de la Fonction  MsgBox (pas de parenthèses nécessaires)'
End Sub
```

</details>

</div>

---

#### Question 6

Configurer les messages contextuels d'entrée et de sortie de la question précédénte, de manière à:
    - ajouter le titre `Saisie année` à le fenêtre de saisie, et paramétrer la valeur 2023 par défaut;
    - ajouter le titre `Taux du livret A`, ne paramétrer qu'un seul bouton `OK`, et ajouter une icône d'alerte.


##### Correction

<div style="border-left:solid #17a589 4px;padding-left:10px; ">
<details>

```php
'Question 6'
Sub question_6()
Dim a As Integer
Dim taux As Double
Dim rep As Double

a = InputBox("Veuillez saisir une année comprise entre 2017 et 2023", "Saisie année", 2023)

If a < 2017 Or a > 2023 Then
    Exit Sub
End If
If a >= 2017 And a <= 2019 Then
        taux = 0.0075
    ElseIf a = 2020 Or a = 2021 Then
        taux = 0.005
    ElseIf a = 2022 Then
        taux = 0.02
    ElseIf a = 2023 Then
        taux = 0.03
End If
MsgBox "Le taux de l'année " & a & " est de: " & Chr(10) & 100 * taux & " %", vbOKOnly + vbExclamation, "Taux du livert A"
'Il s agit encore d un MsgBox en tant que fonction'
End Sub
```

</details>

</div>


---

#### Question 7

Refaire la question précédente en mettant en oeuvre l'instruction `Application.InputBox`, afin de contrôler la saisie d'un nombre entier.

Paramétrer également le messsage de sortie d'écran avec un bouton `Oui`, `Non` et `Annuler`, permettant de récupérer la réponse de l'utilisateur à la question `Etes-vous satisfait?`.

Si la réponse est `Oui`, le message `Bien` s'affichera. Si la réponse est `Non`, le message `Dommage` s'affichera. Si la réponse est `Annuler`, le message `Vous n'avez pas répondu` s'affichera.

##### Correction

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

Pour contrôler la saisie d'un nombre entier, il faut utiliser la fonction `Int` qui renvoie la partie entière d'un nombre flottant.

Pour récupérer la réponse de l'utilisateur à la question `Etes-vous satisfait?`, il faut utiliser la fonction `MsgBox` qui renvoie la valeur de la réponse de l'utilisateur.

Une solution possible est la suivante:

<details>

```php
'Question 7'
Sub question_7()
Dim a As Integer
Dim taux As Double
Dim rep As Double
a = Application.InputBox("Veuillez saisir une année comprise entre 2017 et 2023", "Saisie année", 2023, Type:=1)
If a - Int(a) <> 0 Then  'Int() convertit un flottant en entier. On peut aussi utiliser Fix()'
  MsgBox "Vous n'avez pas saisi un nombre entier"
  Exit Sub
End If

If a >= 2017 And a <= 2019 Then
        taux = 0.0075
    ElseIf a = 2020 Or a = 2021 Then
        taux = 0.005
    ElseIf a = 2022 Then
        taux = 0.02
    ElseIf a = 2023 Then
        taux = 0.03
End If
rep = MsgBox("Le taux de l'année " & a & " est de: " & Chr(10) & 100 * taux & " %", vbYesNoCancel + vbExclamation, "Taux du livert A")
'Il s agit ici de la méthode MsgBox (procédure)= les parenthèses sont obligatoires, ainsi que son affectation à une variable'
If rep = vbYes Then
    MsgBox "Bien"
ElseIf rep = vbNo Then
    MsgBox "Dommage"
Else
    MsgBox "Vous n’avez pas répondu"
End If
End Sub
```

</details>

</div>

---

## Exercice 2

Un placement à intérêt composés se calcule par la formule:

$$C_n = C_0 \times (1 + i)^n$$

avec $C_0$ le capital placé à la date 0 et $C_n$ la valeur acquise par ce capital après $n$ périodes aux taux d'intérêts $i$ par période.

#### Question 1

Ecrire une fonction prenant en paramètres le capital, le nombre de périodes et le taux d'intérêts par période, et renvoyant la valeur du capital acquise.

##### Correction

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

On nous demande d'écrire une fonction prenant en paramètre le capital (un entier), le nombre de périodes (un entier) et le taux d'intérêts par période (un flottant), et renvoyant la valeur du capital acquise (un flottant).

<details>

```php

Function capital(C0 As Double, n As Integer, i As Double) As Double
    capital = C0 * (1 + i) ^ n
End Function
```

</details>

</div>

#### Question 2

Ecrire une précédure affichant le capital acquis dans la cellule **D2** du tableau ci-dessous:

| | A | B | C | D |
| :---: | :---: | :---: | :---: | :---: |
| 1 | Capital | Périodes | Intérêts | Capital acquis |
| 2 | 100 000,00 | 2 | 3% |  |

##### Correction

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

<details>

```php

Sub question_2()
    Dim C0 As Double
    Dim n As Integer
    Dim i As Double
    Dim Cn As Double
    C0 = Range("A2").Value
    n = Range("B2").Value
    i = Range("C2").Value
    Cn = capital(C0, n, i)
    Range("D2").Value = Cn
End Sub
```

</details>

</div>

---

## Exercice 3

#### Question 1

Ecrire un programme qui permet de saisir une note comprise entre 0 et 20 (on affichera un message d'erreur si celle-ci n'est pas comprise entre 0 et 20).

##### Correction

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

<details>

```php

Sub saisir_note()
    Dim note As Double
    note = InputBox("Veuillez saisir une note comprise entre 0 et 20", "Saisie note")
    If note < 0 Or note > 20 Then
        MsgBox "La note saisie n'est pas comprise entre 0 et 20"
        Exit Sub
    End If
End Sub
```

</details>

</div>

#### Question 2

Ajouter au programme précédent l'affichage de la mention relative à la note obtenue (on affichera `Ajourné(e)`si la note est strictement inférieure à 10).

##### Correction

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

<details>

```php

Sub saisir_note()
    Dim note As Double
    Dim mention As String
    note = InputBox("Veuillez saisir une note comprise entre 0 et 20", "Saisie note")
    If note < 0 Or note > 20 Then
        MsgBox "La note saisie n'est pas comprise entre 0 et 20"
        Exit Sub
    End If
    If note < 10 Then
        mention = "Ajourné(e)"
    ElseIf note < 12 Then
        mention = "Passable"
    ElseIf note < 14 Then
        mention = "Assez bien"
    ElseIf note < 16 Then
        mention = "Bien"
    ElseIf note < 18 Then
        mention = "Très bien"
    Else
        mention = "Excellent"
    End If
    MsgBox "La mention est: " & mention
End Sub
```

</details>

</div>

---

## Exercice 4

Une agence de location de véhicules décide d'automatiser le calcul du prix facturé à ces clients. Ecrire une fonction `Location de véhicules` prenant en paramètres d'entrée le kilométrage, le nombre de jours et la catégorie du véhicule, et renvoyant le prix de la location, sachant que :

- Si le véhicule est loué plys de 30 jours, alors le prix sera calculé par la formule: $75 \times \text{Jour}
- Sinon:
    - Si la catégorie est `luxe`, alors la formule est: $80 \times \text{Jour} + 0.2 \times \text{Km}
    - Si la catégorie est `berline`, alors la formule est: $60 \times \text{Jour} + 0.2 \times \text{Km}
    - Pour toutes les autres catégories, la formule est: $70 \times \text{Jour} + 0.15 \times \text{Km}

##### Correction

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

<details>

```php

Function Location_de_vehicules(Km As Double, Jour As Integer, Categorie As String) As Double
    If Jour > 30 Then
        Location_de_vehicules = 75 * Jour
    Else
        If Categorie = "luxe" Then
            Location_de_vehicules = 80 * Jour + 0.2 * Km
        ElseIf Categorie = "berline" Then
            Location_de_vehicules = 60 * Jour + 0.2 * Km
        Else
            Location_de_vehicules = 70 * Jour + 0.15 * Km
        End If
    End If
End Function
```

</details>

</div>

---

## Exercice 5

Ecrire un programme qui permet à son utilisateur de saisir une valeur entière et qui, en retour, lui indique si cette valeur est un nombre premier ou non.

##### Correction

<div style="border-left:solid #17a589 4px;padding-left:10px; ">

<details>

```php

Sub nombre_premier()
    Dim n As Integer
    Dim i As Integer
    Dim est_premier As Boolean
    n = InputBox("Veuillez saisir un nombre entier", "Saisie nombre")
    est_premier = True
    For i = 2 To n - 1
        If n Mod i = 0 Then
            est_premier = False
            Exit For
        End If
    Next i
    If est_premier Then
        MsgBox "Le nombre est premier"
    Else
        MsgBox "Le nombre n'est pas premier"
    End If
End Sub
```

</details>

</div>
