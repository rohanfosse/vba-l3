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

Pensez à activer la case Développeur dans les options d'Excel. Pour cela, allez dans le menu **Fichier**  puis **Options**  et cliquez sur **Personnaliser le ruban**. Dans la fenêtre qui s'ouvre, cliquez sur **Développeur**  dans la liste de gauche et cochez la case **Afficher la barre de développeur**.

Lorsque c'est fait, vous pouvez maintenant ouvrir la fenêtre de VBA en cliquant sur le bouton **Développeur**  dans la barre d'outils et en cliquant sur **Visual Basic**.

---

## Les variables <a name="-les-variables"></a>

Une **variable** en informatique est une sorte de **boîte** dans laquelle on peut mettre une valeur qui peut être utilisée ou modifiée par le programme. Elle permet au programme de **stocker** des informations temporairement ou de façon permanente.

<div class="exemple">
Imaginez que vous êtes en train de préparer votre démenagement et que vous disposez de carton pour ranger vos affaires. Vous pouvez écrire sur chaque carton ce qu'il contient (par exemple un carton **vaisselle** qui va contenir toute votre vaisselle). Dans ce carton, vous pouvez mettre des assiettes, des verres, des bols, etc.

Un autre carton **livre** peut contenir tous vos livres. Les cartons représentent les variables.

Vous pouvez remplir les cartons et les vider. De la même manière, vous pouvez **modifier** les valeurs des variables.

Chaque type de carton ne peut contenir qu'un type d'objet. Par exemple, un carton **vaisselle** ne peut contenir que des assiettes, des verres, des bols, etc. Il ne peut pas contenir des livres. C'est exactement la même chose pour une variable en informatique. Une variable ne peut contenir qu'un seul **type** de valeur à la fois.

En informatique, on appelle cela les **types de variables**.
</div>

En **VBA**, on peut déclarer des variables de plusieurs types :

- **Integer**  : entier
- **Long**  : entier long
- **Single**  : nombre à virgule flottante
- **Double**  : nombre à virgule flottante
- **String**  : chaîne de caractères
- **Boolean**  : booléen (vrai ou faux)

Il existe aussi un type **Variant**  qui permet de stocker n'importe quel type de variable.

Pour déclarer une variable, on utilise la syntaxe suivante :

```vb
Dim nom_variable As type_variable
```

<div class="exemple">

Pour déclarer une variable **a**  de type **Integer**, on utilise la syntaxe suivante :

```vb
Dim a As Integer
```

Si nous reprenons l'analogie avec les cartons, on peut dire que la variable **a**  est un carton qui contient un entier. On ne peut donc y stocker qu'un seul entier, mais on peut le modifier à tout moment.

</div>


Pour déclarer plusieurs variables de même type, on peut utiliser la syntaxe suivante :

```vb
Dim a As Integer, b As Integer, c As Integer
```

Attention, la syntaxe suivante n'est pas valide:

```vb
Dim a, b, c As Integer
```

En effet, on défini ici les variables **a**  et **b**  comme étant de type **Variant**  et la variable **c**  comme étant de type **Integer**.

### Valeur d'une variable

Pour changer la valeur d'une variable, on utilise la syntaxe suivante :

```vb
nom_variable = valeur
```

<div class="exemple">

Si l'on souhaite déclarer une variable **a**  de type **Integer**  et lui donner la valeur **5** :

```vb
Dim a As Integer
a = 5
```

</div>

Pour afficher la valeur d'une variable, on utilise la syntaxe suivante :

```vb
MsgBox nom_variable
```

<div class="exemple">

Si l'on souhaite afficher la valeur de la variable **a**  déclarée précédemment :

```vb
MsgBox a
```

</div>

Si l'on souhaite changer la valeur d'une variable en fonction de sa valeur actuelle, on peut le faire de la manière suivante :

```vb
nom_variable = nom_variable + valeur
```

Dans notre cas, le **nom_variable** de gauche correspond à la **nouvelle** valeur de la variable.

<div class="exemple">

Si l'on souhaite augmenter la valeur de la variable **a**  de **1**  (on dit **incrémenter** en informatique):

```vb
Dim a As Integer
a = 5
a = a + 1 'On incrémente la valeur de a de 1'
MsgBox a ' affiche 6'
```

</div>


---

## Les blocs de code <a name="-les-blocs-de-code"></a>

Un bloc de code est un ensemble d'instructions qui sont exécutées les unes après les autres.

Pour définir un bloc de code, on utilise la syntaxe suivante :

```vb
Bloc nom()
    ' instructions '
End Bloc
```

Il est important de noter que si l'on ouvre un bloc de code, il faut le fermer avec **End Bloc**.

Les différents blocs de code sont :

- **Function**  : pour définir une fonction
- **Sub**  : pour définir une procédure
- **If**  : pour définir une condition
- **Select Case**  : pour définir une condition switch

---

## Les fonctions <a name="-les-fonctions"></a>

En informatique, une fonction est une séquence d'instructions qui effectuent une tâche spécifique et qui peuvent être appelées et réutilisées plusieurs fois dans un programme.

Une fonction peut être comparée à une recette de cuisine. Tout comme une recette de cuisine décrit les étapes à suivre pour préparer un plat, une fonction décrit les étapes à suivre pour accomplir une tâche particulière dans un programme. Les fonctions prennent souvent des entrées, appelées **arguments**, et peuvent renvoyer une sortie, appelée **valeur de retour**.

Par exemple, une fonction **additionner** pourrait prendre deux nombres comme arguments et renvoyer la somme de ces nombres comme valeur de retour. Cette fonction pourrait être appelée **plusieurs fois** dans un programme pour effectuer des opérations d'addition différentes.

Pour déclarer une fonction, on utilise la syntaxe suivante :

```vb
Function nom_fonction() As type_variable
    ' instructions '
    nom_fonction = valeur
End Function
```

Comme la fonction retourne une valeur, il est important de préciser le type de la valeur retournée.

La ligne **nom_fonction = valeur**  permet de renvoyer une valeur à la fonction.

Il est important de noter que le type de la fonction et de la valeur retournée doivent être **identiques**.

<div class="exemple">

Par exemple, pour déclarer une fonction **retourner_a**  qui retourne la lettre "a", on utilise la syntaxe suivante :

```vb
Function retourner_a() As String
    retourner_a = "a"
End Function
```

</div>

Dans notre exemple, la fonction **retourner_a**  retourne une valeur de type **String**.

Si jamais nous souhaitons retourner une valeur de type **Integer**, par exemple **1**, il faut modifier la fonction comme suit :

<div class="exemple">

```vb
Function retourner_a() As Integer
    retourner_a = 1
End Function
```

</div>

Une fonction peut avoir plusieurs paramètres. Pour déclarer une fonction avec plusieurs paramètres, on utilise la syntaxe suivante :

```vb
Function nom_fonction(paramètre1 As type_variable, paramètre2 As type_variable) As type_variable
    ' instructions '
    nom_fonction = valeur
End Function
```

Par exemple, si nous souhaitons déclarer une fonction **aire_rectangle**  qui retourne l'aire d'un rectangle, on utilise la syntaxe suivante :

<div class="exemple">

```vb
Function aire_rectangle(longueur As Integer, largeur As Integer) As Integer
    aire_rectangle = longueur * largeur
End Function
```

</div>

---

## Les procédures <a name="-les-procedures"></a>

Une procédure est une fonction qui ne renvoie pas de valeur.
Pour déclarer une procédure, on utilise la syntaxe suivante :

```vb
Sub nom_procédure()
    ' instructions '
End Sub
```

<div class="exemple">

Pour déclarer la procédure **afficher_message**, on utilise la syntaxe suivante :

```vb
Sub afficher_message()
    MsgBox "Mon message"
End Sub
```

</div>

Pour appeler une procédure, il suffit de cliquer sur le code de la procédure et d'appuyer sur **F5**.

---

## Les conditions <a name="-les-conditions"></a>

En informatique, une condition est un test logique qui permet de prendre une décision en fonction d'une situation donnée.

On peut comparer une condition à une bifurcation dans un chemin : si une condition est remplie, on prend un chemin spécifique ; sinon, on prend un autre chemin.

Par exemple, imaginons que vous écrivez un programme pour vérifier si une personne est autorisée à acheter de l'alcool en fonction de son âge. Vous pouvez utiliser une condition pour tester si l'âge de la personne est supérieur ou égal à 18 ans. Si la condition est vraie, alors la personne est autorisée à acheter de l'alcool ; sinon, elle ne l'est pas.

En informatique, on utilise souvent des instructions conditionnelles, comme "si" ou "if" en anglais, pour mettre en place des conditions. Ces instructions permettent au programme de décider quoi faire en fonction de la situation rencontrée.

Pour définir une condition, on utilise la syntaxe suivante :

```vb
If condition Then
    ' instructions '
End If
```

Si jamais la condition n'est pas vérifiée, le bloc d'instructions n'est pas exécuté. Pour exécuter un bloc d'instructions si la condition n'est pas vérifiée, on utilise le mot clé **else**.

Dans ce cas, on utilise la syntaxe suivante :

```vb
If condition Then
    ' instructions '
Else
    ' instructions '
End If
```

Le **else if**  est une condition supplémentaire qui est exécutée si la condition précédente n'est pas vérifiée. Il peut se traduire par **sinon si**  en français.

Pour définir une condition avec plusieurs **else if**, on utilise la syntaxe suivante :

```vb
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

Prenons l'exemple d'une fonction **appreciation**  qui retourne une appréciation différente suivant une note donné en paramètre.

On peut définir cette fonction comme suit :

```vb
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

- La fonction **appreciation**  retourne une valeur de type **String**.
- La fonction **appreciation**  a un paramètre **note**  de type **Integer**.
- Si jamais une condition n'est pas vérifiée, alors on passe à la condition suivante.
- Si jamais aucune condition n'est vérifiée, alors on exécute le **else**.

Pour appeler cette fonction, on peut définir la procédure suivante :

```vb
Sub afficher_appreciation()
    MsgBox appreciation(15)
End Sub
```

De la même façon que précédement, il est à noter que la valeur du paramètre **note**  peut être n'importe quelle valeur de type **Integer**.

</div>

---

## Les opérateurs logiques <a name="-les-operateurs-logiques"></a>

Les opérateurs logiques permettent de comparer des valeurs entre elles. Les opérateurs logiques sont les suivants :

- **=**  : égal à
- **<>**  : différent de
- **>**  : supérieur à
- **<**  : inférieur à
- **>=**  : supérieur ou égal à
- **<=**  : inférieur ou égal à

Par exemple, pour comparer deux variables **a**  et **b**, on utilise la syntaxe suivante :

```vb
If a = b Then
    ' instructions '
End If
```

Si l'on souhaite faire plusieurs comparaisons à la suite, on peut utiliser les opérateurs logiques suivants :

- **And**  : et
- **Or**  : ou
- **Not**  : non

<div class="exemple">
Par exemple, pour comparer deux variables **a**  et **b**, on utilise la syntaxe suivante :

```vb
If a = b And a > 0 Then
    ' instructions '
End If
```

</div>

---

## Select Case <a name="-select-case"></a>

Pour définir une condition avec plusieurs **else if**, il existe une autre méthode utilisant la syntaxe suivante :

```vb
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

Les deux façons de définir une condition sont **équivalentes**, le **Select Case**  est simplement une méthode plus concise.

Les mots clés autorisés dans un **Case**  sont les suivants :

- **Is**  : égal à
- **Is Not**  : différent de
- **>**  : supérieur à
- **<**  : inférieur à
- **>=**  : supérieur ou égal à
- **<=**  : inférieur ou égal à
- **To**  : entre

<div class="exemple">
Si nous reprenons l'exemple de la fonction **appreciation**  précédente, on peut définir cette fonction comme suit :

```vb
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

Il faut noter que puisque j'ai changé le nom de la fonction en **appreciation_select**  (pour ne pas confondre avec la fonction **appreciation**  précédente), je dois modifier l'affectation **appreciation_select**  (à la place de **appreciation`).

Une façon de traduire ce code en français serait :

```vb
Je sélectionne la variable note.
- Si je suis dans le cas où la note est comprise entre 0 et 9, alors appréciation = "ajournée"
- Si je suis dans le cas où la note est comprise entre 10 et 11, alors appréciation = "passable"
- Si je suis dans le cas où la note est comprise entre 12 et 13, alors appréciation = "assez bien"
- Si je suis dans le cas où la note est comprise entre 14 et 15, alors appréciation = "bien"
- Si je suis dans le cas où la note est comprise entre 16 et 17, alors appréciation = "très bien"
- Sinon, appréciation = "excellent"
```

De la même façon que pour les **Else If**, on peut définir la procédure suivante pour appeler la fonction:

```vb
Sub afficher_appreciation()
    MsgBox appreciation(15)
End Sub
```
</div>
---

## Les fenêtres prédéfinies (InputBox et MsgBox) <a name="-les-fenetres-predefinies"></a>

Il existe plusieurs fenêtres prédéfinies en VBA.

#### Fenêtre de saisie de texte (_InputBox_)

La saisie de texte se fait avec la fenêtre **InputBox**.

<div class="exemple">
Par exemple, pour afficher la fenêtre **InputBox**  avec le message **Entrez un nombre**  et stocker le résultat dans la variable **nombre**, on utilise la syntaxe suivante :

```vb
Dim nombre As Integer ' Déclaration de la variable nombre de type Integer '
nombre = InputBox("Entrez un nombre")
```

A la suite de ça, la variable **nombre**  contient la valeur saisit par l'utilisateur.

</div>

Si l'on souhaite maintenant afficher la même fenêtre mais en changeant le titre par **Mon titre**, on utilise la syntaxe suivante :

<div class="exemple">

```vb
Dim nombre As Integer ' Déclaration de la variable nombre de type Integer'
Dim titre As String ' Déclaration de la variable titre de type String'

titre = "Mon titre"
nombre = InputBox("Entrez un nombre", titre)
```

</div>

Enfin, si l'on souhaite en plus que la valeur par défaut soit **1**, on utilise la syntaxe suivante :

<div class="exemple">

```vb
Dim nombre As Integer ' Déclaration de la variable nombre de type Integer'
Dim titre As String ' Déclaration de la variable titre de type String'
Dim valeur_par_defaut As Integer ' Déclaration de la variable defaut de type Integer'

titre = "Mon titre"
valeur_par_defaut = 1
nombre = InputBox("Entrez un nombre", titre, valeur_par_defaut)
```

</div>


Dans le cas où l'utilisateur ne saisit rien, la variable **nombre**  contient la valeur **1**.

---

#### Fenêtre d'affichage de message (_MsgBox_)

L'affichage d'un message se fait avec la fenêtre **MsgBox**.

<div class="exemple">

Par exemple, pour afficher la fenêtre **MsgBox**  avec le message "a", on utilise la syntaxe suivante :

```vb
MsgBox("Mon message")
```

</div>

Si l'on souhaite afficher un message comportant une variable **v**, on utilise la syntaxe suivante :

<div class="exemple"

```vb
MsgBox("Mon message" & v)
```

</div>

Il est possible de modifier les boutons affichés dans la fenêtre **MsgBox**  en ajoutant des valeurs à la fin de la fonction.

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

Par exemple, pour afficher la fenêtre **MsgBox**  avec le message "Mon message" et le bouton "OK", on utilise la syntaxe suivante :

```vb
resultat = MsgBox("Mon message", vbOKOnly)
```

</div>

La variable **resultat**  contient la valeur **1**  si l'utilisateur clique sur le bouton "OK".

Pour afficher la fenêtre **MsgBox**  avec le message "Mon message", un bouton "OK' et le titre "Titre", on utilise la syntaxe suivante :

<div class="exemple">

```vb
resultat = MsgBox("Mon message", vbOKOnly, "Titre")
```

</div>

Pour afficher la fenêtre **MsgBox**  avec le message "Mon message", le titre "Titre" et les boutons "Oui" et "Non", on utilise la syntaxe suivante :

<div class="exemple">

```vb
resultat = MsgBox("Mon Message", vbYesNo, "Titre")
```

</div>

On stocke la réponse de l'utilisateur dans la variable **resultat**. Si l'utilisateur clique sur le bouton "Oui", la variable **resultat**  contient la valeur **6**. Si l'utilisateur clique sur le bouton "Non", la variable **resultat**  contient la valeur **7**.


<div class="exemple">

Par exemple, si l'on souhaite poser une question à l'utilisateur et afficher un message différent en fonction de sa réponse, on peut utiliser le faire de la façon suivante :

```vb
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
Ecrire une fonction **perimetre**  calculant le perimètre d'un cercle prenant en paramètre un entier correspondant à son rayon. Le résultat sera un réel.
La valeur 3.14 sera utilisée pour la constante pi.
On affichera le résultat à l'aide d'une procédure **afficher_perimetre**.

##### Solution possible

<details>
{% highlight vb %}
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
Ecrire une fonction **calculer_moyenne**  calculant la moyenne de 3 notes qui seront données en paramètre de la fonction. Le résultat sera un réel.

##### Solution possible

<details>
{% highlight vb %}
Function calculer_moyenne(note1 As Integer, note2 As Integer, note3 As Integer) As Single
calculer_moyenne = (note1 + note2 + note3) / 3
End Function
{% endhighlight %}
</details>

</div>


#### Exercice 3 <a name="exercice-3"></a>

<div class="exemple_blue">
Ecrire une fonction **calculer_somme**  calculant la somme de 2 entiers que l'utilisateur saisira à l'aide de deux fenêtres **InputBox**. Si l'utilisateur ne saisit pas de valeur, la valeur par défaut sera 0.

##### Solution possible

<details>
{% highlight vb %}
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