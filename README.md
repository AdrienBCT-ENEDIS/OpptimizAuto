# Opptimizer Automatic

- Automatisation de l'outil Opptimiz dévelopé par le groupe EDF.

# Usage

./opptimizAuto.ps1 source 

source    [répertoire à optimiser]

ou

Lancez "exe.bat" puis choisissez le dossier à optimiser.


# Détails

Le programme parcours le répertoire ainsi que ses sous-répertoires pour récupérer tous les fichiers ".pptx" et les optimiser. Par défaut, le programme écrase les fichiers et utilise le mode de compression "Intermédiaire".
Si vous voulez mettre le mode de compression "Maximum" il suffit d'enlever "{RIGHT}" (ligne 38) et si vous voulez conserver le fichier d'origine il faut rajouter à la ligne 41 un "{TAB}" à la suite des deux autres.

Assurez-vous que le ruban de raccourci est bien activé sur PowerPoint avant de lancer le programme (CTRL+F1)

Lorsque vous lancez le programme, ne touchez à rien sur votre machine et attendez la fin de l'exécution du programme.