#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
from ssigie_helpers.AD_Helper import AD_Helper
import re
import sys
from datetime import datetime
from ssigie_utils.Mail_Utils import is_valid_email, send_email
import os

def main():
    parser=argparse.ArgumentParser(description="Utilitaire d'extraction des utilisateurs et groupes AD")
    parser.add_argument('--ad', choices=['GIECB', 'CB.LOCAL'], required=True, help="Spécifie l'AD à requêter: 'GIECB' ou 'CB.LOCAL'")
    parser.add_argument('--extract', choices=['users', 'members', 'groups'], required=True, help="Spécifie quoi exporter: 'users' pour les utilisateurs, 'groups' pour les groupes, 'members' pour les membres de groupes.")
    parser.add_argument('--mail', type=str, help='Adresse e-mail pour export (optionnel)')
    parser.add_argument('--output', type=str, help='Répertoire de sortie')
    parser.add_argument('--conf', type=str, help='Fichier de configuration')

    args=parser.parse_args()

    if args.mail:
        if not is_valid_email(args.mail):
            print(f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}][ERROR] {args.mail} is not a valid adress.")
            return 1

    dirtarget="./Outputs"
    if args.output:
      dirtarget=args.output
    if not os.path.exists(dirtarget):
      os.makedirs(dirtarget)

    conftarget="./config.ini"
    if args.conf:
      conftarget=args.conf
    # Chargement de AD_Helper - classe permettant de requêter l'AD selon les besoins
    ad_helper=AD_Helper(conftarget)

    if args.extract == 'users':
        # Extraction uniquement des utilisateurs
        output_file= f"%s/Export_AD_Users_{args.ad}_{datetime.now().strftime('%Y-%m-%d')}.xlsx" % (dirtarget)
        ad_helper.get_users(args.ad, output_file)
        if args.mail:
            mail_subject=f"Export AD {args.ad} du {datetime.now().strftime('%d/%m/%Y')}"
            mail_body=f"Bonjour,\n\nCi-joint un export des utilisateurs de l'AD {args.ad}.\n\nCe message est automatique."

    elif args.extract == 'members':
        # Extraction des groupes et des utilisateurs
        output_file= f"%s/Export_AD_GroupMembers_{args.ad}_{datetime.now().strftime('%Y-%m-%d')}.xlsx" % (dirtarget)
        ad_helper.get_user_groups(args.ad, output_file)
        if args.mail:
            mail_subject=f"Export des membres des groupes AD {args.ad} du {datetime.now().strftime('%d/%m/%Y')}"
            mail_body=f"Bonjour,\n\nCi-joint un export des groupes et de leurs membres de l'AD {args.ad}.\n\nCe message est automatique."

    elif args.extract == 'groups':
        # Extraction des groupes et des utilisateurs
        output_file=f"%s/Export_AD_Groups_{args.ad}_{datetime.now().strftime('%Y-%m-%d')}.xlsx" % (dirtarget)
        ad_helper.get_groups(args.ad, output_file)
        if args.mail:
            mail_subject=f"Export des groupes AD {args.ad} du {datetime.now().strftime('%d/%m/%Y')}"
            mail_body=f"Bonjour,\n\nCi-joint un export des groupes de l'AD {args.ad}.\n\nCe message est automatique."
            
    if args.mail:
        # Envoi du mail si demandé
        send_email(
            smtp_server='localhost',    # Local MTA (postfix configuré pour forward sur MTA fourni par SIG)
            sender_email='vmw-gcb-029@cartes-bancaires.com',
            receiver_email=args.mail,
            subject=mail_subject,
            body=mail_body,
            attachment_paths=[output_file]
        )

if __name__ == '__main__':
    main()
