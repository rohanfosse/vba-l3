<p style="text-align:left;">
    [Retour au sommaire](../../README.md)
</p>

<div style="text-align:center;">
# Correction du TD 4 - Structures Itératives
</div>

---

## Exercice 1

Ecrire une procédure demandant à l'utilisateur de saisir un nombre $n$, puis affichant $n$ fois le message "Message i" (avec i prenant les valeurs de 1 à $n$).

Donner trois variantes de cette procédure: l'une avec une boucle `For`, l'autre avec une boucle `While` et la dernière avec une boucle `Do While`.

#### Correction

<details>

```php
Sub exo1()
    Dim n As Integer
    Dim i As Integer
    
    n = InputBox("Entrez un nombre", "Nombre")
    
    For i = 1 To n ' Variante For'
        MsgBox "Message " & i
    Next i
    
    i = 1
    While i <= n 'Variante While'
        MsgBox "Message " & i
        i = i + 1
    Wend
    
    i = 1
    Do While i <= n 'Variante Do While'
        MsgBox "Message " & i
        i = i + 1
    Loop
End Sub
```

</details>

## Exercice 2

Ecrire une fonction permettant de calculer le prix actualisé d'un produit, après $n$ années écoulées avec un taux d'inflation à 5% (taux annuel).

*NB: la fonction prendra en paramètres d'entrée le prix initial p et le nombre d'années n.*

Obtiendrait-on la même valeur en appliquant le taux $n \times 5\%$ au prix initial?

#### Correction

<details>

```php
Function prixActualise(p As Double, n As Integer) As Double
    Dim i As Integer
    Dim prix As Double
    
    prix = p
    
    For i = 1 To n
        prix = prix * 1.05
    Next i
    
    prixActualise = prix
End Function

Sub afficherPrixActualise()
    Dim p As Double
    Dim n As Integer
    
    p = InputBox("Entrez le prix initial", "Prix")
    n = InputBox("Entrez le nombre d'années", "Années")
    
    MsgBox "Le prix actualisé est de " & prixActualise(p, n)
End Sub
```

</details>

## Exercice 3

On considère un placement à intérêt composé, avec capitalisation annuelle des interêts, en notant:

- $C$ le montant du caputal placé (en euros);
- $n$ la durée du placement (en années);
- $i$ le taux d'intérêt.

#### Question 1

Ecrire un programme qui affiche sur une feuille de calcul (comme ci-dessous), les montants annuels capitalisés d'un placement à intérêt composé dont les caractéristiques (capital placé, durée et taux d'intérêt) sont demandées par des messages contextuels.

| Année | Montant |
|-------|---------|
| 1     | 1000    |
| 2     | 1050    |
| 3     | 1102.5  |
| 4     | 1157.63 |
| -     | -       |

#### Correction

<details>

```php
Sub exo3q1()
    Dim C As Double
    Dim n As Integer
    Dim i As Double
    Dim j As Integer
    
    C = InputBox("Entrez le capital placé", "Capital")
    n = InputBox("Entrez la durée du placement", "Durée")
    i = InputBox("Entrez le taux d'intérêt", "Taux")
    
    Cells(1, 1) = "Année"
    Cells(1, 2) = "Montant"
    
    For j = 1 To n
        Cells(j + 1, 1) = j
        Cells(j + 1, 2) = C * (1 + i) ^ j
    Next j
End Sub
```

</details>

##### Question 2

Reprendre la question précédente, en mettant en oeuvre la formule de récurrence du calcul des montatns annuels capitalisés: $C_{n} = C_{n-1} \times (1+i)$.

#### Correction

<details>

```php
Sub exo3q2()
    Dim C As Double
    Dim n As Integer
    Dim i As Double
    Dim j As Integer
    
    C = InputBox("Entrez le capital placé", "Capital")
    n = InputBox("Entrez la durée du placement", "Durée")
    i = InputBox("Entrez le taux d'intérêt", "Taux")
    
    Cells(1, 1) = "Année"
    Cells(1, 2) = "Montant"
    
    For j = 1 To n
        Cells(j + 1, 1) = j
        Cells(j + 1, 2) = C * (1 + i) ^ j
    Next j
End Sub
```

</details>

## Exercice 4

Ecrire un programme qui affiche tous les nombres parfaits compris entre 2 et 10000.

Un nombre parfait est un enteier égale à la somme de ses diviseurs, lui exclu.

Par exemple, 28 est un nombre parfait car $1+2+4+7+14=28$.

#### Correction

<details>

```php
Sub exo4()
    Dim i As Integer
    Dim j As Integer
    Dim somme As Integer
    
    For i = 2 To 10000
        somme = 0
        For j = 1 To i - 1
            If i Mod j = 0 Then
                somme = somme + j
            End If
        Next j
        If somme = i Then
            MsgBox i
        End If
    Next i
End Sub
```

</details>