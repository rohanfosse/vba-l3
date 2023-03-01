# Programmation VBA

Vous trouverez ici des notes de cours relatives au cours de VBA de l'Université de Bordeaux Montaigne.
Vous pouvez trouver les slides [ici](https://fad4.u-bordeaux.fr/pluginfile.php/2050621/mod_resource/content/1/3_S%C3%A9ance_3_VBA%20pour%20Excel_L3_Part1.pdf)

## Avant de commencer

Pensez à activer la case Développeur dans les options d'Excel. Pour cela, allez dans le menu `Fichier` puis `Options` et cliquez sur `Personnaliser le ruban`. Dans la fenêtre qui s'ouvre, cliquez sur `Développeur` dans la liste de gauche et cochez la case `Afficher la barre de développeur`.

## Les bases

### Les variables

Une variable est un espace mémoire qui permet de stocker une valeur. En VBA, on peut déclarer des variables de plusieurs types :

* `Integer` : entier
* `Long` : entier long
* `Single` : nombre à virgule flottante
* `Double` : nombre à virgule flottante
* `String` : chaîne de caractères
* `Boolean` : booléen (vrai ou faux)

Pour déclarer une variable, on utilise la syntaxe suivante :

    Dim nom_variable As type_variable

Par exemple, pour déclarer une variable `a` de type `Integer`, on utilise la syntaxe suivante :

    Dim a As Integer

Pour déclarer plusieurs variables du même type, on peut utiliser la syntaxe suivante :

    Dim a, b, c As Integer

### Les procédures

Une procédure est une fonction qui ne renvoie pas de valeur. Pour déclarer une procédure, on utilise la syntaxe suivante :

    Sub nom_procédure()
        ' instructions
    End Sub

Par exemple, pour déclarer une procédure `afficher_a`, on utilise la syntaxe suivante :

    Sub afficher_a()
        MsgBox "a"
    End Sub

Pour appeler une procédure, il suffit de cliquer sur le code de la procédure et d'appuyer sur `F5`.

### Les fonctions

Une fonction est une procédure qui renvoie une valeur. Pour déclarer une fonction, on utilise la syntaxe suivante :

    Function nom_fonction() As type_variable
        ' instructions
        nom_fonction = valeur
    End Function

Par exemple, pour déclarer une fonction `retourner_a` qui retourne la lettre "a", on utilise la syntaxe suivante :

    Function retourner_a() As String
        retourner_a = "a"
    End Function

Pour afficher cette fonction, on peut définir la procédure suivante :

    Sub afficher_a()
        MsgBox retourner_a()
    End Sub

Une fonction peut avoir plusieurs paramètres. Pour déclarer une fonction avec plusieurs paramètres, on utilise la syntaxe suivante :

    Function nom_fonction(paramètre1 As type_variable, paramètre2 As type_variable) As type_variable
        ' instructions
        nom_fonction = valeur
    End Function



### Les conditions

Pour définir une condition, on utilise la syntaxe suivante :

    If condition Then
        ' instructions
    End If

Pour définir une condition avec un `else`, on utilise la syntaxe suivante :

    If condition Then
        ' instructions
    Else
        ' instructions
    End If

Pour définir une condition avec plusieurs `else if`, on utilise la syntaxe suivante :

    If condition Then
        ' instructions
    ElseIf condition Then
        ' instructions
    ElseIf condition Then
        ' instructions
    Else
        ' instructions
    End If

Par exemple, pour déclarer une fonction `retourner_a_b` qui retourne la lettre "a" si le paramètre `a` vaut 1, et la lettre "b" sinon, on utilise la syntaxe suivante :

        Function retourner_a_b(a As Integer) As String
            If a = 1 Then
                retourner_a_b = "a"
            Else
                retourner_a_b = "b"
            End If
        End Function

Pour appeler cette fonction, on peut définir la procédure suivante :

    Sub afficher_a_b()
        MsgBox retourner_a_b(1)
    End Sub

Il est à noter que la valeur du paramètre `a` peut être n'importe quelle valeur de type `Integer`.

### Select Case

Pour définir une condition avec plusieurs `else if`, une autre méthode est d'utiliser la syntaxe suivante :

    Select Case variable
        Case valeur1
            ' instructions
        Case valeur2
            ' instructions
        Case valeur3
            ' instructions
        Case Else
            ' instructions
    End Select

Par exemple, pour déclarer une fonction `retourner_a_b` qui retourne la lettre "a" si le paramètre `a` vaut 1, et la lettre "b" sinon, on utilise la syntaxe suivante :

        Function retourner_a_b_case(a As Integer) As String
            Select Case a
                Case 1
                    retourner_a_b_case = "a"
                Case Else
                    retourner_a_b_case = "b"
            End Select
        End Function

Pour appeler cette fonction, on peut définir la procédure suivante :

    Sub afficher_a_b_case()
        MsgBox retourner_a_b_case(1)
    End Sub

### Les fenêtres prédéfinies

Il existe plusieurs fenêtres prédéfinies en VBA.

La saisie de texte se fait avec la fenêtre `InputBox`. Par exemple, pour afficher la fenêtre `InputBox` avec le message "Entrez un nombre", on utilise la syntaxe suivante :

    InputBox("Entrez un nombre")

Pour afficher la fenêtre `InputBox` avec le message "Entrez un nombre" et le titre "Nombre", on utilise la syntaxe suivante :

    InputBox("Entrez un nombre", "Nombre")

Pour afficher la fenêtre `InputBox` avec le message "Entrez un nombre", le titre "Nombre" et la valeur par défaut "1", on utilise la syntaxe suivante :

    InputBox("Entrez un nombre", "Nombre", "1")

L'affichage d'un message se fait avec la fenêtre `MsgBox`. Par exemple, pour afficher la fenêtre `MsgBox` avec le message "a", on utilise la syntaxe suivante :

    MsgBox "a"

Pour afficher la fenêtre `MsgBox` avec le message "a" et le titre "Message", on utilise la syntaxe suivante :

    MsgBox "a", vbInformation, "Message"

Pour afficher la fenêtre `MsgBox` avec le message "a", le titre "Message" et le bouton "OK", on utilise la syntaxe suivante :

    MsgBox "a", vbInformation + vbOKOnly, "Message"

Pour afficher la fenêtre `MsgBox` avec le message "a", le titre "Message" et les boutons "Oui" et "Non", on utilise la syntaxe suivante :

    MsgBox "a", vbInformation + vbYesNo, "Message"


