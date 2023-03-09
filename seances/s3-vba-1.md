<p style="text-align:left;">
    [Retour au sommaire](../README.md)
    <span style="float:right;">
        [Séance 4 - VBA: Structures itératives et communication avec Excel](s4-vba-2.md)
    </span>
</p>

<div style="text-align:center;">
# Séance 3 - VBA: Fonctions et procédures, variables, tests et boucles
</div>
---

## Avant de commencer

Pensez à activer la case Développeur dans les options d'Excel. Pour cela, allez dans le menu `Fichier` puis `Options` et cliquez sur `Personnaliser le ruban`. Dans la fenêtre qui s'ouvre, cliquez sur `Développeur` dans la liste de gauche et cochez la case `Afficher la barre de développeur`.

---

## Les variables <a name="-les-variables"></a>

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

<div class="exemple">

Pour déclarer une variable `a` de type `Integer`, on utilise la syntaxe suivante :

```php
Dim a As Integer
```

</div>

Pour déclarer plusieurs variables de même type, on peut utiliser la syntaxe suivante :

```php
Dim a As Integer, b As Integer, c As Integer
```

Attention, la syntaxe suivante n'est pas valide:

```php
Dim a, b, c As Integer
```

En effet, on défini ici les variables `a` et `b` comme étant de type `Variant` et la variable `c` comme étant de type `Integer`.

---

## Les blocs de code <a name="-les-blocs-de-code"></a>

Un bloc de code est un ensemble d'instructions qui sont exécutées les unes après les autres.

Pour définir un bloc de code, on utilise la syntaxe suivante :

```php
Bloc nom()
    ' instructions '
End Bloc
```

Il est important de noter que si l'on ouvre un bloc de code, il faut le fermer avec `End Bloc`.

Les différents blocs de code sont :

- `Sub` : pour définir une procédure
- `Function` : pour définir une fonction
- `If` : pour définir une condition
- `For` : pour définir une boucle for
- `While` : pour définir une boucle while
- `Do` : pour définir une boucle do while
- `Select Case` : pour définir une condition switch

---

## Les procédures <a name="-les-procedures"></a>

Une procédure est une fonction qui ne renvoie pas de valeur.
Pour déclarer une procédure, on utilise la syntaxe suivante :

```php
Sub nom_procédure()
    ' instructions '
End Sub
```

<div class="exemple">

Pour déclarer la procédure `afficher_message`, on utilise la syntaxe suivante :

```php
Sub afficher_message()
    MsgBox "Mon message"
End Sub
```

</div>

Pour appeler une procédure, il suffit de cliquer sur le code de la procédure et d'appuyer sur `F5`.

---

## Les fonctions <a name="-les-fonctions"></a>

Une fonction est une procédure qui renvoie une valeur. Pour déclarer une fonction, on utilise la syntaxe suivante :

```php
Function nom_fonction() As type_variable
    ' instructions '
    nom_fonction = valeur
End Function
```

Comme la fonction retourne une valeur, il est important de préciser le type de la valeur retournée.

La ligne `nom_fonction = valeur` permet de renvoyer une valeur à la fonction.

Il est important de noter que le type de la fonction et de la valeur retournée doivent être **identiques**.

<div class="exemple">

Par exemple, pour déclarer une fonction `retourner_a` qui retourne la lettre "a", on utilise la syntaxe suivante :

```php
Function retourner_a() As String
    retourner_a = "a"
End Function
```

</div>

Dans notre exemple, la fonction `retourner_a` retourne une valeur de type `String`.

Si jamais nous souhaitons retourner une valeur de type `Integer`, par exemple `1`, il faut modifier la fonction comme suit :

<div class="exemple">

```php
Function retourner_a() As Integer
    retourner_a = 1
End Function
```

</div>
Pour afficher cette fonction, on peut définir la procédure suivante :

<div class="exemple">

```php
Sub afficher_a()
    MsgBox retourner_a()
End Sub
```

</div>

Une fonction peut avoir plusieurs paramètres. Pour déclarer une fonction avec plusieurs paramètres, on utilise la syntaxe suivante :

```php
Function nom_fonction(paramètre1 As type_variable, paramètre2 As type_variable) As type_variable
    ' instructions '
    nom_fonction = valeur
End Function
```

Par exemple, si nous souhaitons déclarer une fonction `aire_rectangle` qui retourne l'aire d'un rectangle, on utilise la syntaxe suivante :

<div class="exemple">

```php
Function aire_rectangle(longueur As Integer, largeur As Integer) As Integer
    aire_rectangle = longueur * largeur
End Function
```

</div>

Pour appeler cette fonction, on peut définir la procédure suivante :

<div class="exemple">
```php
Sub afficher_aire_rectangle()
    MsgBox aire_rectangle(10, 5)
End Sub
```

</div>


Il est important de noter que dans l'appel de la fonction `aire_rectangle`, on ne met pas les noms des paramètres mais on donne directement les **valeurs** des paramètres.

Ainsi, une autre façon d'appeler cette procédure serait par exemple :

<div class="exemple">

```php
Sub afficher_aire_rectangle()
    MsgBox aire_rectangle(8, 10)
End Sub
```

</div>

---

## Les conditions <a name="-les-conditions"></a>

Une insctruction conditonnelle est une instruction qui permet d'exécuter un bloc de code si une condition est vérifiée.

Pour définir une condition, on utilise la syntaxe suivante :

```php
If condition Then
    ' instructions '
End If
```

Si jamais la condition n'est pas vérifiée, le bloc d'instructions n'est pas exécuté. Pour exécuter un bloc d'instructions si la condition n'est pas vérifiée, on utilise le mot clé `else`.

Dans ce cas, on utilise la syntaxe suivante :

```php
If condition Then
    ' instructions '
Else
    ' instructions '
End If
```

Le `else if` est une condition supplémentaire qui est exécutée si la condition précédente n'est pas vérifiée. Il peut se traduire par `sinon si` en français.

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

<div class="exemple">

Prenons l'exemple d'une fonction `appreciation` qui retourne une appréciation différente suivant une note donné en paramètre.

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

</div>

---

## Les opérateurs logiques <a name="-les-operateurs-logiques"></a>

Les opérateurs logiques permettent de comparer des valeurs entre elles. Les opérateurs logiques sont les suivants :

- `=` : égal à
- `<>` : différent de
- `>` : supérieur à
- `<` : inférieur à
- `>=` : supérieur ou égal à
- `<=` : inférieur ou égal à

Par exemple, pour comparer deux variables `a` et `b`, on utilise la syntaxe suivante :

```php
If a = b Then
    ' instructions '
End If
```

Si l'on souhaite faire plusieurs comparaisons à la suite, on peut utiliser les opérateurs logiques suivants :

- `And` : et
- `Or` : ou
- `Not` : non

<div class="exemple">
Par exemple, pour comparer deux variables `a` et `b`, on utilise la syntaxe suivante :

```php
If a = b And a > 0 Then
    ' instructions '
End If
```

</div>

---

## Select Case <a name="-select-case"></a>

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

Les deux façons de définir une condition sont **équivalentes**, le `Select Case` est simplement une méthode plus concise.

Les mots clés autorisés dans un `Case` sont les suivants :

- `Is` : égal à
- `Is Not` : différent de
- `>` : supérieur à
- `<` : inférieur à
- `>=` : supérieur ou égal à
- `<=` : inférieur ou égal à
- `To` : entre

<div class="exemple">
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

Il faut noter que puisque j'ai changé le nom de la fonction en `appreciation_select` (pour ne pas confondre avec la fonction `appreciation` précédente), je dois modifier l'affectation `appreciation_select` (à la place de `appreciation`).

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
</div>
---

## Les fenêtres prédéfinies (InputBox et MsgBox) <a name="-les-fenetres-predefinies"></a>

Il existe plusieurs fenêtres prédéfinies en VBA.

#### Fenêtre de saisie de texte (_InputBox_)

La saisie de texte se fait avec la fenêtre `InputBox`.

<div class="exemple">
Par exemple, pour afficher la fenêtre `InputBox` avec le message `Entrez un nombre` et stocker le résultat dans la variable `nombre`, on utilise la syntaxe suivante :

```php
Dim nombre As Integer ' Déclaration de la variable nombre de type Integer '
nombre = InputBox("Entrez un nombre")
```

A la suite de ça, la variable `nombre` contient la valeur saisit par l'utilisateur.

</div>

Si l'on souhaite maintenant afficher la même fenêtre mais en changeant le titre par `Mon titre`, on utilise la syntaxe suivante :

<div class="exemple">

```php
Dim nombre As Integer ' Déclaration de la variable nombre de type Integer'
Dim titre As String ' Déclaration de la variable titre de type String'

titre = "Mon titre"
nombre = InputBox("Entrez un nombre", titre)
```

</div>

Enfin, si l'on souhaite en plus que la valeur par défaut soit `1`, on utilise la syntaxe suivante :

<div class="exemple">

```php
Dim nombre As Integer ' Déclaration de la variable nombre de type Integer'
Dim titre As String ' Déclaration de la variable titre de type String'
Dim valeur_par_defaut As Integer ' Déclaration de la variable defaut de type Integer'

titre = "Mon titre"
valeur_par_defaut = 1
nombre = InputBox("Entrez un nombre", titre, valeur_par_defaut)
```

</div>


Dans le cas où l'utilisateur ne saisit rien, la variable `nombre` contient la valeur `1`.

---

#### Fenêtre d'affichage de message (_MsgBox_)

L'affichage d'un message se fait avec la fenêtre `MsgBox`.

<div class="exemple">

Par exemple, pour afficher la fenêtre `MsgBox` avec le message "a", on utilise la syntaxe suivante :

```php
MsgBox("Mon message")
```

</div>

Si l'on souhaite afficher un message comportant une variable **v**, on utilise la syntaxe suivante :

<div class="exemple"

```php
MsgBox("Mon message" & v)
```

</div>

Il est possible de modifier les boutons affichés dans la fenêtre `MsgBox` en ajoutant des valeurs à la fin de la fonction.

Les noms, valeurs et significations pour les principaux boutons peuvent être trouvés dans le tableau suivant :

| Nom | Valeur | Signification |
| :--- | :--- | :--- |
| vbOKOnly | 0 | Affiche uniquement le bouton "OK" |
| vbOKCancel | 1 | Affiche les boutons "OK" et "Annuler" |
| vbAbortRetryIgnore | 2 | Affiche les boutons "Annuler", "Réessayer" et "Ignorer" |
| vbYesNoCancel | 3 | Affiche les boutons "Oui", "Non" et "Annuler" |
| vbYesNo | 4 | Affiche les boutons "Oui" et "Non" |
| vbRetryCancel | 5 | Affiche les boutons "Réessayer" et "Annuler" |

<div class="exemple">

Par exemple, pour afficher la fenêtre `MsgBox` avec le message "Mon message" et le bouton "OK", on utilise la syntaxe suivante :

```php
resultat = MsgBox("Mon message", vbOKOnly)
```

</div>

La variable `resultat` contient la valeur `1` si l'utilisateur clique sur le bouton "OK".

Pour afficher la fenêtre `MsgBox` avec le message "Mon message", un bouton "OK' et le titre "Titre", on utilise la syntaxe suivante :

<div class="exemple">

```php
resultat = MsgBox("Mon message", vbOKOnly, "Titre")
```

</div>

Pour afficher la fenêtre `MsgBox` avec le message "Mon message", le titre "Titre" et les boutons "Oui" et "Non", on utilise la syntaxe suivante :

<div class="exemple">

```php
resultat = MsgBox("Mon Message", vbYesNo, "Titre")
```

</div>

On stocke la réponse de l'utilisateur dans la variable `resultat`. Si l'utilisateur clique sur le bouton "Oui", la variable `resultat` contient la valeur `6`. Si l'utilisateur clique sur le bouton "Non", la variable `resultat` contient la valeur `7`.


<div class="exemple">

Par exemple, si l'on souhaite poser une question à l'utilisateur et afficher un message différent en fonction de sa réponse, on peut utiliser le faire de la façon suivante :

```php
Sub afficher_message()
    Dim resultat As Integer
    resultat = MsgBox("Voulez-vous continuer ?", vbYesNo, "Titre")
    If resultat = 6 Then
        MsgBox("Vous avez cliqué sur Oui")
    Else
        MsgBox("Vous avez cliqué sur Non")
    End If
End Sub
```

</div>

---

## Exercices Corrigés <a name="-exercices-corriges"></a>

Si vous souhaitez vous entrainer, voici quelques exercices corrigés.

#### Exercice 1 <a name="exercice-1"></a>

<div class="exemple_blue">
Ecrire une fonction `perimetre` calculant le perimètre d'un cercle prenant en paramètre un entier correspondant à son rayon. Le résultat sera un réel.
La valeur 3.14 sera utilisée pour la constante pi.
On affichera le résultat à l'aide d'une procédure `afficher_perimetre`.

##### Solution possible

<details>
{% highlight php %}
Function perimetre(rayon As Integer) As Single
    perimetre = 2 * 3.14 * rayon
End Function

Sub afficher_perimetre()
    MsgBox perimetre(1)
End Sub
{% endhighlight %}
</details>
</div>

#### Exercice 2 <a name="exercice-2"></a>

<div class="exemple_blue">
Ecrire une fonction `calculer_moyenne` calculant la moyenne de 3 notes qui seront données en paramètre de la fonction. Le résultat sera un réel.

##### Solution possible

<details>
{% highlight php %}
Function calculer_moyenne(note1 As Integer, note2 As Integer, note3 As Integer) As Single
calculer_moyenne = (note1 + note2 + note3) / 3
End Function
{% endhighlight %}
</details>

</div>


#### Exercice 3 <a name="exercice-3"></a>

<div class="exemple_blue">
Ecrire une fonction `calculer_somme` calculant la somme de 2 entiers que l'utilisateur saisira à l'aide de deux fenêtres `InputBox`. Si l'utilisateur ne saisit pas de valeur, la valeur par défaut sera 0.

##### Solution possible

<details>
{% highlight php %}
Function calculer_somme() As Integer
Dim nombre1 As Integer
Dim nombre2 As Integer

nombre1 = InputBox("Entrez un nombre", "Nombre 1", 0)
nombre2 = InputBox("Entrez un nombre", "Nombre 2", 0)

calculer_somme = nombre1 + nombre2
End Function
{% endhighlight %}
</details>
</div>

---

## Correction TD3 <a name="correction-td3"></a>

Vous pouvez trouver la correction du TD3 [ici](exercices/s3-td1.md)