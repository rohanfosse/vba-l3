# Programmation VBA

Vous trouverez ici des notes de cours relatives au cours de VBA de l'Université de Bordeaux Montaigne.
Vous pouvez trouver le cours sur cette page [moodle](https://fad4.u-bordeaux.fr/enrol/index.php?id=17638)

---

## Table des matières

<!--@import "[TOC]" {cmd="toc" depthFrom=2 depthTo=3 orderedList=false}-->

<!-- code_chunk_output -->

- [Avant de commencer](#-avant-de-commencer)
- [Séance 3](#-seance-3)
  - [Les variables](#-les-variables)
  - [Les procédures](#-les-procédures)
  - [Les fonctions](#-les-fonctions)
  - [Les conditions](#-les-conditions)
  - [Select Case](#-select-case)
  - [Les fenêtres prédéfinies](#-les-fenêtres-prédéfinies)
  - [Exemples supplémentaires ](#-exemples-supplémentaires-a-nameexemplesa)
  - [Exercices Corrigés](#-exercices-corrigés)

<!-- /code_chunk_output -->

---

## Avant de commencer <a name="-avant-de-commencer"></a>

Pensez à activer la case Développeur dans les options d'Excel. Pour cela, allez dans le menu `Fichier` puis `Options` et cliquez sur `Personnaliser le ruban`. Dans la fenêtre qui s'ouvre, cliquez sur `Développeur` dans la liste de gauche et cochez la case `Afficher la barre de développeur`.

---

## Séance 3 <a name="-seance-3"></a>

### Les variables <a name="-les-variables"></a>

Une variable est un espace mémoire qui permet de stocker une valeur. En `VBA`, on peut déclarer des variables de plusieurs types :

- `Integer` : entier
- `Long` : entier long
- `Single` : nombre à virgule flottante
- `Double` : nombre à virgule flottante
- `String` : chaîne de caractères
- `Boolean` : booléen (vrai ou faux)

Il existe aussi un type `Variant` qui permet de stocker n'importe quel type de variable.

Pour déclarer une variable, on utilise la syntaxe suivante :

```php
Dim nom_variable As type_variable
```

Par exemple, pour déclarer une variable `a` de type `Integer`, on utilise la syntaxe suivante :

```php
Dim a As Integer
```

Pour déclarer plusieurs variables du même type, on peut utiliser la syntaxe suivante :

```php
Dim a As Integer, b As Integer, c As Integer
```

Attention, la syntaxe suivante n'est pas valide:

```php
Dim a, b, c As Integer
```

En effet, cette syntaxe défini les variables `a` et `b` comme étant de type `Variant` et la variable `c` comme étant de type `Integer`.

---

### Les procédures <a name="-les-procédures"></a>

Une procédure est une fonction qui ne renvoie pas de valeur. Pour déclarer une procédure, on utilise la syntaxe suivante :

```php
Sub nom_procédure()
    ' instructions '
End Sub
```

Par exemple, pour déclarer une procédure `afficher_a`, on utilise la syntaxe suivante :

```php
Sub afficher_a()
    MsgBox "a"
End Sub
```

Pour appeler une procédure, il suffit de cliquer sur le code de la procédure et d'appuyer sur `F5`.

---

### Les fonctions <a name="-les-fonctions"></a>

Une fonction est une procédure qui renvoie une valeur. Pour déclarer une fonction, on utilise la syntaxe suivante :

```php
Function nom_fonction() As type_variable
    ' instructions '
    nom_fonction = valeur
End Function
```

La ligne `nom_fonction = valeur` permet de renvoyer une valeur à la fonction.

Il est important de noter que la valeur retournée doit être du même type que le type de la fonction.

Par exemple, pour déclarer une fonction `retourner_a` qui retourne la lettre "a", on utilise la syntaxe suivante :

```php
Function retourner_a() As String
    retourner_a = "a"
End Function
```

Dans notre exemple, la fonction `retourner_a` retourne une valeur de type `String`.

Si jamais nous souhaitons retourner une valeur de type `Integer`, par exemple 1, il faudra modifier la fonction comme suit :

```php
Function retourner_a() As Integer
    retourner_a = 1
End Function
```

Pour afficher cette fonction, on peut définir la procédure suivante :

```php
Sub afficher_a()
    MsgBox retourner_a()
End Sub
```

Une fonction peut avoir plusieurs paramètres. Pour déclarer une fonction avec plusieurs paramètres, on utilise la syntaxe suivante :

```php
Function nom_fonction(paramètre1 As type_variable, paramètre2 As type_variable) As type_variable
    ' instructions '
    nom_fonction = valeur
End Function
```

Par exemple, si nous souhaitons déclarer une fonction `aire_rectangle` qui retourne l'aire d'un rectangle, on utilise la syntaxe suivante :

```php
Function aire_rectangle(longueur As Integer, largeur As Integer) As Integer
    aire_rectangle = longueur * largeur
End Function
```

Pour appeler cette fonction, on peut définir la procédure suivante :

```php
Sub afficher_aire_rectangle()
    MsgBox aire_rectangle(10, 5)
End Sub
```

Il est important de noter que dans l'appel de la fonction `aire_rectangle`, on ne met pas les noms des paramètres mais on donne directement les **valeurs** des paramètres.

Ainsi, une autre façon d'appeler cette procédure serait par exemple :

```php
Sub afficher_aire_rectangle()
    MsgBox aire_rectangle(8, 10)
End Sub
```

---

### Les conditions <a name="-les-conditions"></a>

Pour définir une condition, on utilise la syntaxe suivante :

```php
If condition Then
    ' instructions '
End If
```

Pour définir une condition avec un `else`, on utilise la syntaxe suivante :

```php
If condition Then
    ' instructions '
Else
    ' instructions '
End If
```

Le `else` est exécuté si la condition n'est pas vérifiée. Il peut se traduire par `sinon` en français.

Pour définir une condition avec plusieurs `else if`, on utilise la syntaxe suivante :

```php
If condition Then
    ' instructions '
ElseIf condition Then
    ' instructions '
ElseIf condition Then
    ' instructions '
Else
    ' instructions '
End If
```

Le `else if` est une condition supplémentaire qui est exécutée si la condition précédente n'est pas vérifiée. Il peut se traduire par `sinon si` en français.

Prenons l'exemple d'une fonction `appreciation` qui retourne une appréciation différante suivant une note donné en paramètre.

On peut définir cette fonction comme suit :

```php
Function appreciation(note As Integer) As String
    If note < 10 Then
        appreciation = "ajourné"
    ElseIf note < 12 Then
        appreciation = "passable"
    ElseIf note < 14 Then
        appreciation = "assez bien"
    ElseIf note < 16 Then
        appreciation = "bien"
    ElseIf note < 18 Then
        appreciation = "très bien"
    Else
        appreciation = "excellent"
    End If
End Function
```

Dans ce code, plusieurs points sont à noter :

- La fonction `appreciation` retourne une valeur de type `String`.
- La fonction `appreciation` a un paramètre `note` de type `Integer`.
- Si jamais une condition n'est pas vérifiée, alors on passe à la condition suivante.
- Si jamais aucune condition n'est vérifiée, alors on exécute le `else`.

Pour appeler cette fonction, on peut définir la procédure suivante :

```php
Sub afficher_appreciation()
    MsgBox appreciation(15)
End Sub
```

De la même façon que précédement, il est à noter que la valeur du paramètre `note` peut être n'importe quelle valeur de type `Integer`.

---

### Select Case <a name="-select-case"></a>

Pour définir une condition avec plusieurs `else if`, il existe une autre méthode utilisant la syntaxe suivante :

```php
    Select Case variable
        Case valeur1
            ' instructions '
        Case valeur2
            ' instructions '
        Case valeur3
            ' instructions '
        Case Else
            ' instructions '
    End Select
```

Les deux façons de définir une condition sont **_équivalentes_**, le `Select Case` est simplement une méthode plus concise.

Si nous reprenons l'exemple de la fonction `appreciation` précédente, on peut définir cette fonction comme suit :

```php
Function appreciation_select(note As Integer) As String
    Select Case note
        Case 0 To 9
            appreciation_select = "ajourné"
        Case 10 To 11
            appreciation_select = "passable"
        Case 12 To 13
            appreciation_select = "assez bien"
        Case 14 To 15
            appreciation_select = "bien"
        Case 16 To 17
            appreciation_select = "très bien"
        Case Else
            appreciation_select = "excellent"
    End Select
End Function
```

Il faut noter que puisque j'ai changé le nom de la fonction en `appreciation_select` (pour ne pas confondre avec la fonction `appreciation` précédente), je dois modifier l'affectation `appreciation_select` à la place de `appreciation`.

Une façon de traduire ce code en français serait :

```php
Je sélectionne la variable note.
- Si je suis dans le cas où la note est comprise entre 0 et 9, alors appréciation = "ajournée"
- Si je suis dans le cas où la note est comprise entre 10 et 11, alors appréciation = "passable"
- Si je suis dans le cas où la note est comprise entre 12 et 13, alors appréciation = "assez bien"
- Si je suis dans le cas où la note est comprise entre 14 et 15, alors appréciation = "bien"
- Si je suis dans le cas où la note est comprise entre 16 et 17, alors appréciation = "très bien"
- Sinon, appréciation = "excellent"
```

De la même façon que pour les `Else If`, on peut définir la procédure suivante pour appeler la fonction:

```php
Sub afficher_appreciation()
    MsgBox appreciation(15)
End Sub
```

---

### Les fenêtres prédéfinies <a name="-les-fenêtres-prédéfinies"></a>

Il existe plusieurs fenêtres prédéfinies en VBA.

#### Fenêtre de saisie de texte (_InputBox_)

La saisie de texte se fait avec la fenêtre `InputBox`. Par exemple, pour afficher la fenêtre `InputBox` avec le message `Entrez un nombre` et stocker le résultat dans la variable `nombre`, on utilise la syntaxe suivante :

```php
Dim nombre As Integer ' Déclaration de la variable nombre de type Integer '
nombre = InputBox("Entrez un nombre")
```

A la suite de ça, la variable `nombre` contient la valeur saisie par l'utilisateur.

Si l'on souhaite maintenant afficher la même fenêtre mais en changeant le titre par `Mon titre`, on utilise la syntaxe suivante :

```php
Dim nombre As Integer ' Déclaration de la variable nombre de type Integer'
Dim titre As String ' Déclaration de la variable titre de type String'

titre = "Mon titre"
nombre = InputBox("Entrez un nombre", titre)
```

Enfin, si l'on souhaite en plus que la valeur par défaut soit `1`, on utilise la syntaxe suivante :

```php
Dim nombre As Integer ' Déclaration de la variable nombre de type Integer'
Dim titre As String ' Déclaration de la variable titre de type String'
Dim defaut As Integer ' Déclaration de la variable defaut de type Integer'

titre = "Mon titre"
valeur_par_defaut = 1
nombre = InputBox("Entrez un nombre", titre, valeur_par_defaut)
```

Dans le cas où l'utilisateur ne saisit rien, la variable `nombre` contient la valeur `1`.

Pour d'autres exemples, voir la section [Exemples supplémentaires](#exemples).

---

#### Fenêtre d'affichage de message (_MsgBox_)

L'affichage d'un message se fait avec la fenêtre `MsgBox`. Par exemple, pour afficher la fenêtre `MsgBox` avec le message "a", on utilise la syntaxe suivante :

```php
MsgBox "a"
```

Pour afficher la fenêtre `MsgBox` avec le message "a" et le titre "Message", on utilise la syntaxe suivante :

```php
MsgBox "a", vbInformation, "Message"
```

Pour afficher la fenêtre `MsgBox` avec le message "a", le titre "Message" et le bouton "OK", on utilise la syntaxe suivante :

```php
MsgBox "a", vbInformation + vbOKOnly, "Message"
```

Pour afficher la fenêtre `MsgBox` avec le message "a", le titre "Message" et les boutons "Oui" et "Non", on utilise la syntaxe suivante :

```php
MsgBox "a", vbInformation + vbYesNo, "Message"
```

---

### Exemples supplémentaires <a name="exemples"></a>

#### Exemple de fonction utilisant une condition (vérification de notes)

#### En utilisant une structure `if`

```php
Function verifier_note(note As Integer) As String
    If note < 10 Then
        verifier_note = "Insuffisant"
    ElseIf note < 12 Then
        verifier_note = "Passable"
    ElseIf note < 14 Then
        verifier_note = "Assez bien"
    ElseIf note < 16 Then
        verifier_note = "Bien"
    ElseIf note < 20 Then
        verifier_note = "Très bien"
    Else
        verifier_note = "Erreur"
    End If
End Function
```

#### En utilisant une structure `select case`

```php
Function verifier_note_case(note As Integer) As String
    Select Case note
        Case 0 To 9
            verifier_note_case = "Insuffisant"
        Case 10 To 11
            verifier_note_case = "Passable"
        Case 12 To 13
            verifier_note_case = "Assez bien"
        Case 14 To 15
            verifier_note_case = "Bien"
        Case 16 To 19
            verifier_note_case = "Très bien"
        Case Else
            verifier_note_case = "Erreur"
    End Select
End Function
```

---

### Exercices Corrigés <a name="exercices-corrigés"></a>

#### Exercice 1 <a name="exercice-1"></a>

Ecrire une fonction `perimetre` calculant le perimètre d'un cercle prenant en paramètre un entier correspondant à son rayon. Le résultat sera un réel.
La valeur 3.14 sera utilisée pour la constante pi.
On affichera le résultat à l'aide d'une procédure `afficher_perimetre`.


<details>
<summary>Solution</summary>

{% highlight php %}
Function perimetre(rayon As Integer) As Single
perimetre = 2 _ 3.14 _ rayon
End Function

Sub afficher_perimetre()
MsgBox perimetre(1)
End Sub
{% endhighlight %}

</details>
