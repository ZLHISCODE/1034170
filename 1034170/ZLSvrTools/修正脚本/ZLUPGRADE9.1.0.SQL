--�޸Ĺ��߱�
--ע�⣺��zltools�û���¼
--Connect zltools/zltools
ALTER TABLE zlFunctions ADD CONSTRAINT zlFunctions_UQ_���� Unique (ϵͳ,������)
/
Alter Table zlFuncPars Add (������ VARCHAR2(40),��ʽ Number(2))
/
Alter Table zlFuncPars Modify ���� VARCHAR2(30)
/

DELETE FROM  zlSvrTools WHERE ���='05'
/
insert into zlSvrTools(���,�ϼ�,����,���,˵��) values ('05',Null,'ר���','R',Null)
/
insert into zlSvrTools(���,�ϼ�,����,���,˵��) values ('0501','05','��������','R',Null)
/
insert into zlSvrTools(���,�ϼ�,����,���,˵��) values ('0502','05','��������','F',Null)
/


CREATE OR REPLACE Function zlSpellCode(v_InStr IN VARCHAR2)
    RETURN VARCHAR2
    IS v_Spell VARCHAR2(10);
    v_Input VARCHAR2(1000);
    v_BitChar VARCHAR2(2);
    v_BitNum INTEGER;
    v_ChrNum INTEGER;
    v_StdStr VARCHAR2(50):='�Ų���귢������-��������Ŷž��Ȼ������-����ѹ��';
    v_CharA VARCHAR2(2000):='߹��H�����X�����t�H�c���\���������ٌ���r�P�o�a�������Y��O��c�g��������@���@���t؁�B���l���E�J�������֒�������q�������O���֓��������';
    v_CharB VARCHAR2(2000):='�����^���������R�T�������Zڕ���E��������ٔ�v�C�n���������k����[��k�ߙ�D���r�^����������_�����R�dم�d����E�s�U�t���l�mؐ����f�K�����R݅�^�������G�S��Q����ݙ�a����J�M��G�a�P���q�S�sݩذ����������؄םߛ���������������[�]����C�P�����G�s�����޵�I�ۋ�@��������@�S�{�K�L�x�F�E�z��������߄�Q��߅�b�c���H���u��������������r�g�Y�l�m�p�q׃�������R���������[րٙ�S���j�k�l�n�s�Օ�l���B�M������ߓ�������h�e�f�����E�\����������W������v����u���V���@�m�h����\�G�@�Q���R������K�����c�D�N�m�n�P���}��������L��������c�J�K��߲�G�L�Q����b�Y�^����������������������������';
    v_CharC VARCHAR2(2000):='�����nؔ�P�{�����o���֍���]�I�ى�������[ܳ������������������\��x���ۂ���d������O٭�{���i������S�a�������������K׋������~����p֝�P�U׀�����]����������_����K�����L�M��������L�l����������k�o�����n���}���C�{֚��܇������ފ�J�����Հ�o�nލ����ܕ�\��R���mڒ���{ٕ���f��Y�Z��׏�p�����W�l��X�d�d�p�Kة���J�������\��������������v�|�A���c�J�[���d�P�W�m�gܯ���t�s���r�X��߳��ތ������L�M��o�چ���u���������ی���|ٱ����O�l���P�~ׇ׉�h�{�c��،�iۻ�a�n�I�z�r���R�X�����s��ء����X�s�U�zՑ������������������ݎ����A�i�E�J���������N�m�qݐ�j���ݻ�c�T����ك�w�����u�O��Q�}�z���p�q�w���e�����e����i��@�y�o�\�]ڝ�n��������W�����p�z�{ց����ݏ����������������K�u�q���y���A���ߥ���f������J��~�������xߗ�Z���u���i�������z�������H��S�e���������������������������������������������';
    v_CharD VARCHAR2(2000):='������pޅއ������Q�_�Q�����]�J�N�^����߾�a�߰�����܍ܤܖ�J�D�\�l��ې�O�n�^���l�G�����F���^����Q��}ٜ�K���[�hו��������T�ځ�W��߶���������O�I�Z��܄����u�O��������Q�����������h�����~�L�C��ݶ���{�E��p�Mصڮۡ�s�������B���W�ޞ���K����f�d���r�B�y��ۆ���������c���������������J�M�������m����y�H��{��S���ܦ���g��Pՙ�����l��������w�����}������V�r���M�A�����[���C���HՉ�����ޓ��������h���K�^�K�Y�H�k�Z�W�L�L�^�`�a�^��`�A������i�L���x�K�G�~���o�b��tטـܶ�|��H�����Y�X��������������B��m���H�����O������ޚ�g�D�q�v��������y�I����r�o܀�D���w�y�F�G��z�����������������������������������';
    v_CharE VARCHAR2(2000):='��ވݭ�e��M�~�P��~�Z�[�F��E�i�����q������ܗ������`�Q�]�����{�O�Iج�@������t��׆�y�|�{�����E�z�[���X�b���W�������s�D߃ڍ٦�@�E����';
    v_CharF VARCHAR2(2000):='�e���y���z��N�c�x�Y�Cެ�������x���x�܏��؜�G����������ړ�[�p�h���������J���w����q�����E�y����������u���M������]���m�p����V���X�k�����r�M�M�R�v������a�����L�����h�S��Qۺ�b�p�K���Sٺ�R�L�P�iو��]ߑ߻���K�������a�~�f�W���A�F����ܽ�������������ۮ��I��ݳ���������E�R�V�O�D�h��q�D�~ݗ�H�v���f��߼���M����o�f������ؓ������xݕ�Vَ�����v���������������';
    v_CharG VARCHAR2(2000):='�٤������m�����p�@�������B�d�W�^�Yؤ�}��|�����������N�v�����s��ߦ���h�������M�l��G��s����غ�z�k���������޻�ھ۬��a������x�m�w�g���ت��ܪ����w����k���Yݑ�s�k�u�P�R���������t��ب��ݢ���f�s���Q������������p���������\�ؕ�C�����h�^���x����xڸ������g��ُ�������L�M�����Y��ݞ���X��ڬ����E�����������]����������d���A����N�o�T�W������ڴ���v�K���P���b݄�]�I�A����؞�k���q���X�}�������_��U���ߞ�F�|���q���h�k�I������{܉�������F���W�Z�iح�����݁���P��֏�������u�����������{�R�J�^���������������������������������';
    v_CharH VARCHAR2(2000):='���x�����V�����E�A���������w�n�J�_�\���F�\���I�d�h�u֛�n��[�ކ�ؘ�@������޶����q�����ڭ�����A������؀�F�����u�i���M�����H�H�[���Y���a�R���Q�L�e�f�S�g�\������a���C޿�U��ܟ�Fްݓ��Z����ݦ�A�v�b�D�p���fޮ�������Zڧ�{�U���\�J������A�\�C�����jܩ�_�����`�c�������ܠ����i֗��������������g��L�E�{���C�K����t�U��������������U��o���_��s�I�j�k�����n֜�f����Ֆՠ�X���b�J��؎גۨ����f�}���a���߀�o�D�I�q�S�X�G���kۼ�������ߧ���Z�d�����S���������B��W���ڇ��u�m�U���e�wڶ�������Y�D�x���������ޒ�D�t�e���������������_�V�dޥ�M�_�M�f�i�T��u�}�w�������Q�F�@ڻ��՟�����x߫ߘ�����X؛�A�f�o����޽�Z�[���������������������������������';
    v_CharJ VARCHAR2(2000):='آߴ����ܸ�����������|���������K����}�u����Z�Y�ي���u�I����^��Q�Z�a�V�W�i�W�A�ؽ٥�B�L�C�������l�ު�e��n݋ۈ��U�g�P�W�n�e�|�}����������������������E�H���H������Ղ�����J۔�a�H�T�����D�V���C�q������������e���j؆����ۣ�P������O��e�]�a���G���Z�����]�����g�y�������b���Z��K�������p�x�t���Y�[���~�d�����������������ֈ�C�r�x���{�v�|�����V����ڙ�f�`��Ր�v�{�`���G�I�T�G�Z�Y�a�b�{������������\�F�v������n���uܴ֘����������o���B�����ٮ������]�]�q����a�K�R��^��ڊ�I���_��A���������ڦڵ��ޗ����]��������m�d�R�O�^����]�����v�T������������\֔�ݣ���M������ف�B�����ݼ���X�L�~������������������i������ޟ�����e�V�K�n�o�R������ޛ�G�y������F���b�N��������J���n�����������ڠ���Շ�g�|�L�~�����`�]�R�z�v����V۞�q�G���A�������F���eڪ��ߚ�������B���e���X����Z��؋�M���L��g�m����N������������۲�h��C�������������bڑ�I���k�f�����ާ�_�`�����Q���H���B�����~�������؏�j�܊�x�z�������z�����}���K�Q�R�U�����������������������������������������������';
    v_CharK VARCHAR2(2000):='�������l��_��������]�����a���|�z�G�a���b�����٩ݨ�|ݝ��R����R�{�_�K��ߒ������`���������D���w������ڐ���V�����W�f�w�������������n��c�~�o��U�L��H�����w�x�I�y�����ߵ�@ޢ�d��ߠܥڜ���p���F٨���ۦ�����������d���wڲߝ���E�Hڿܒܜ�N�\�����������L�A�k�q���Y�k��N����k�ظ�������`��K���i���ۓ�������d�q�^������d���K���H�{�A��������S��A�p�H�T�U��������������������';
    v_CharL VARCHAR2(2000):='��������h�F�J�_�n�B������[�F��n�D���H�����l��ه�m�s�`�������@�E׎�_�|��e�����Y����Oݹ����H�q�Z����L������������������u�L�~���b߷��������E���������D�[�h�Y�m�Fڳ�C�|��P����L�[�G�K���ܨ�k����������؂�����ւ�r޼߆��x���\�P�v���g�~�Z�Gٵ�����N����������ߊ�k߿��������ٳ���\��۪����ݰ�����������W�E�_�t�`���B�b���V�]ׁ�^�Z�u�c���B����ۚ֋�`�`�H����������������b���n����cܮ��݈���u�gՏ�v�y�G�������|�Iْێ��m�������ޤ�R�������ޘ�������V�h�����Q���v�������O�����������O�l���[�������C�U�����C���\�k�`����������_�����ښ�C��q�s����C���o�w���h۹�N�g��`�����t�I�����������v���m�H���y���B�d�s�i���V���ۉ�C�w�f�j�w���������������X�N�����L�[�x�_�T���]�L������s֌�}���V�����U��ߣ������������������_�z�B�|�R�u�u���������u��������T�`��ڀ�j��X��h�j���ۍ�A�G�I�c�n�e��������y�L�����X�r�������F��[���s�x����i݆��MՓ���b�������߉������s���������������i��������������������������������������������������������';
    v_CharM VARCHAR2(2000):='�����j��i����U���K��ݤ�I��۽���u�~�@�A������M�N�����ܬ�������֙�N�����I����؈���ܚ�F���^����������T����Q�|������ݮ�d���������Y�[�B�q��z�V�eڛ�m�i�������T�Y�{�������ޫ�������������s�X��L�����i���Q���D�W�_������i�����������S���J����������ڢ������������k������������|�r���@�M�I�����]���������ؿ���p�f�x����������ق��s�F�����h���w�}������ڤ�p������Q�����և�������N����փք�O�O�������{����������a�����ٰ���w�\�����E��a�[������������f��J���������������������������';
    v_CharN VARCHAR2(2000):='�y����~�����vܘ�y�c���ܵޕ�����ؾ�r����a��Q�y�������T����߭�Qث��������D�t����m�[ګ�H��������G�����\����C�؃�r���F���u٣��b�W�Xދ�������R���D�T��݂ۜ�Tإ���|������B�������������c�W�f�h�R����E���mב�b��D����_�V�H���������o�ٯ���r�s�x�P�a�e��k�����������S�����Q�G�����S����Zہ�������������';
    v_CharO VARCHAR2(2000):='���Mک��k֎��t�{����';
    v_CharP VARCHAR2(2000):='����ٽ݇�����W���A��ۘ�o�Q���G��b�����Q�������������������N����B���k����r���������\�����J�o��ܡ�~�A���i�J�m�s�Cا������w�W�t��Y�����C�B�V�o����ۯ�����u����R���Q���dܱ��������|�aߨ���������G�����@�����X՗�����G՛�������g�h�w�Q����o������د���v�ؚ���l�A����ٷ�Z�Z���Z��݃�G����N�w�k۶�c�����O���H�H���������������T����ٟ�h����E���V�������������������������';
    v_CharQ VARCHAR2(2000):='ހ���V������Ճ�p�[����t�K������������ݽږܙ�H���������D����n�o���Rޭ�a�W�������u�}���G�y����ߌ�ܻ��M����H�M������������ڞ�M�����ܷ�����@�T�`�e���U��e�Ս�w��t�v�c�k�B�R�S�aݡ����ܝ�j�@�Q�E�X�Z�b����������l�c����ٻ��݀���������ۄ��ۖ�j�I�j�����m�������b�z���������^�N����@ډ�E�F�A����������Sڈ�y�X���ڽ�~�V�m�I�N��������o��@����W�z�����s�d�������V���_�c��u�����i�W���X�p�����[�������Ո���m�����������^��jڂ���F�G�p�q���G����ٴޝ������U�������g�M�b�F�����j�A���ڰ�r���o�L�@څ�D�|�������O۾���@ޡ���z���޾����d��Y�xޑ�T������C��z��zڹ�������I�b��m�B�����e�j�E�����j�������|��I�o����������������������������������������������';
    v_CharR VARCHAR2(2000):='�������`�X�j׌������N�v���m���������r��ך�����ܐ���z�~�g��J��w��J�~���������F�g�P����݊����k���qߏ���n���޸���}��p�r�z��������M��ܛ݉ި������J��c�tټ�e�}���U�������������';
    v_CharS VARCHAR2(2000):='���ئ�����l�M�S���|�wِ��L�����D�d�����r��������b���fܣ�����C�Q�m֠�o�O�~��������|������������ߍ���������ܑ��^����������Wڨ���]���۷���b�i٠����W�X������օ�l���p�l����i�Y�}��ۿ�����f�d�h�������s��ڷ�����_�Y���h��ߕ������ן��Ք�T���}����v�j���J�H����|���W�j�����K��ًߟ�\�P���A���O�[����X���P���y�z����ݪ�Y�J�v���Z���R����������������߱���B�Y����m�K՜՞�}�S�u��|�a�����������x�ٿ��ܓ�g����S�\ݔ�_�e�����H���n���t�����_���Q�f�T������X���V���B�p���{�t�U�`�l���j��B�f�h�������������l�p���������j�t���F�����l�J�\�r���D�����������ٹ�~�L�|����ڡ����ݿ�����b��n���������g�}�`�����޴������������i�h�����x�p�_�M�qۑ�T���ݴ������m�U�S���r�w�����w�\���ݥ������{�Z������������t�����a�i����C�����������������������������������';
    v_CharT VARCHAR2(2000):='�������B�D�]�������e���`��w�J��F�O�Y�n�c��ۢ�������U�T޷�����؝�۰��Մ�]�U�t�T؍�v�Z�������g�a���y�����ۏ�M�|�U����o���G����}�Z���h�O�S�����E�����z�N�w��ޏ�����[���cػ��߯ؖ��������߂�`���L�R�e�f���Xڄ���pۇ������}�{�Y�[�����n܃������P�ڌ�L�j�V�p������D�c�l������q�`�t��٬�����p��������x�f���A����q�\���N�@�������F�����ߋ����������F�������F�P�b�c����١�����U�P���n�~������j���B�����^�W��W�Cݱ��T����I���h�����Qރܢ���r؇����o������P�j�k�n�s۝ۃ�����`��Z��ܔ�������رי��٢�u����ސ���������|���D�r���������K�z��ڗ�����������������';
    v_CharW VARCHAR2(2000):='��|��ߜ���c��������ؙܹ�Bߐݸ���������n�l�j��[�sٖ�~�@������s�y��ނ�����������ޱ���g�h�����������f�����`���W���d�S�T���������������Ն�c�l�Q��n�t�]�|���^�M�K�E�A�G���~�Z׈�^�d�nݘݜ���������������Y�j������Z�[�������l�f�O��޳�N��ݫ��b���������}����ڏ�E�w�G�_���u���N����c�M�����~���������������q�^�Rأ�@������A�����F�}����`����H�F�I�F������������';
    v_CharX VARCHAR2(2000):='����ۭ��������ݾ�T�R�������q���O�g�F���O��������a�G�H�l؉�v���T�@���^�^���@������E���v��֐��I���e�@�������|��������L�lے�^�h�ۧ�S�M�����i���V�K�_�]�S�U���B�i�P�y������������Y��ݠ��_���T�pՒ�]�����ݲ����������v�w��r۟�]�N�]����e��t�t�Pݍ�_�y�D�v������������`���U޺�����`�@�����^�}�D�R�G�o���Eܼ�_�l�m����x�����`����K�����}�A���a����z��}�P������������^���X�N���{�y���j�U���q�j�[�M����ߢ���C�P�Hא��������������ޯ�����x�C��^�k�K�a����ߔ�ݷ�d��\ܰ���g��cض܌���_�S�]���D������]�o��tߩ����Nܺכל������������T�x����q����P�V�n�M������נ���q՚֞�z�`�Pڼ����������������[��rޣ܎������������՝�X�M�X�~�z����x����������C������X��K�j�Y���`���z���G���L���p�o���޹����������������S�\��ޙ���d�bަ��R�������������������������';
    v_CharY VARCHAR2(2000):='ѹ����f�E���s���������\������������������������۳���Z��iڥ����ܾ�I�Z��������}����ٲ����۱�D���V�C����y�d�o���f�d���j�k��B�|���z�s��������H�������e�V���e����ٞ܂�z�`�I�Jׅ����ח�V�W���Z����}���g���r������U݌����^�{�u�R�F�I�B������زߺ�^س�������U���b��u��P���{�|�c���_�����������o�r��G���_׊������X�y�U����������v�����]�E�d�w�v�E�v����c�������p�t���b�s�p��ڱ������������ޖ�����O�B�D���z�U�V���{�k֖�Fׂ�@�~�������ޠ�r�i�C��}��T�t߮�����d��߽٫����������������ؗ����[�\�N�cژ�W�z������x���k�o�]����޲��؊�lٓ����s�J�G�������g�h�y�{�O�^�gܲ����ה�~�f��������ܧ���N���������w��۴��z�y�����]�l�������Y�i��L���y�[׍��ط�S�gݺ���A�����a��������v�D�s�L�]��܅���K�W����������������h��A۫�I�e�����G����{ܭ��V���a���t�O�����I�x��J���ٸ���x���k�����~������ݯݵޜ�]�K�[��������j��ݒ�O߈��ݬ���B���٧���ޔ�z�R�T�����}�|ߎ�����خ�������C�����D����~���������՘�k�N��uݛ��~���V�i�}�C��ٶ�������h��Z�o��؅���r�����������C���N���������Aع�h��T�y�`�\���I�[��������r�q�O�u�X��M���N���O܆�c�d���g�S��x�t�����Oߖ���ؒ�J���ܫ�������w�@�x�߇���M�R�h���������܋��ڔ�_�X��������g�S�N���V�S�fٚ�ܿ���y���]�m�����q�E�B�q�y۩��i��\���ل�d�j�r�y��������������������������������������������������������';
    v_CharZ VARCHAR2(2000):='ش��������٪پ����ڣگں������������������ۤۥ۵۸����������ݧ����ީ����������ߡߤߪ߬߸�����������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������';

BEGIN
    if v_Instr is null or length(ltrim(v_InStr))=0 then 
        v_Spell:='';
    else
        v_Input:=upper(v_InStr);
        v_Spell:='';
        for v_BitNum in 1..length(v_Input) loop
            v_BitChar:=substr(v_Input,v_BitNum,1);
            if v_BitChar>='��' AND v_BitChar<='��' then
                for v_ChrNum in 1..length(v_StdStr) loop
                    if substr(v_StdStr,v_ChrNum,1)='-' then
                        null;
                    elsif v_BitChar<substr(v_StdStr,v_ChrNum,1) then
                        v_Spell:=v_Spell||chr(64+v_ChrNum);
                        exit;
                    end if;
                end loop;
                if v_BitChar>='��' then
                    v_Spell:=v_Spell||'Z';
                end if;
            elsif ascii(v_BitChar)<256 then
                v_Spell:=v_Spell||v_BitChar;
            elsif instr('���������������',v_BitChar)>0 then
                v_Spell:=v_Spell||chr(ascii(v_BitChar)-41664);
            elsif instr('���£ãģţƣǣȣɣʣˣ̣ͣΣϣУѣңӣԣգ֣ףأ٣�',v_BitChar)>0 then
                v_Spell:=v_Spell||chr(ascii(v_BitChar)-41856);
            elsif instr('����',v_BitChar)>0 then
                v_Spell:=v_Spell||'A';
            elsif instr('����',v_BitChar)>0 then
                v_Spell:=v_Spell||'B';
            elsif instr('����',v_BitChar)>0 then
                v_Spell:=v_Spell||'G';
            elsif instr(v_CharA,v_BitChar)>0 then
                v_Spell:=v_Spell||'A';
            elsif instr(v_CharB,v_BitChar)>0 then
                v_Spell:=v_Spell||'B';
            elsif instr(v_CharC,v_BitChar)>0 then
                v_Spell:=v_Spell||'C';
            elsif instr(v_CharD,v_BitChar)>0 then
                v_Spell:=v_Spell||'D';
            elsif instr(v_CharE,v_BitChar)>0 then
                v_Spell:=v_Spell||'E';
            elsif instr(v_CharF,v_BitChar)>0 then
                v_Spell:=v_Spell||'F';
            elsif instr(v_CharG,v_BitChar)>0 then
                v_Spell:=v_Spell||'G';
            elsif instr(v_CharH,v_BitChar)>0 then
                v_Spell:=v_Spell||'H';
            elsif instr(v_CharJ,v_BitChar)>0 then
                v_Spell:=v_Spell||'J';
            elsif instr(v_CharK,v_BitChar)>0 then
                v_Spell:=v_Spell||'K';
            elsif instr(v_CharL,v_BitChar)>0 then
                v_Spell:=v_Spell||'L';
            elsif instr(v_CharM,v_BitChar)>0 then
                v_Spell:=v_Spell||'M';
            elsif instr(v_CharN,v_BitChar)>0 then
                v_Spell:=v_Spell||'N';
            elsif instr(v_CharO,v_BitChar)>0 then
                v_Spell:=v_Spell||'O';
            elsif instr(v_CharP,v_BitChar)>0 then
                v_Spell:=v_Spell||'P';
            elsif instr(v_CharQ,v_BitChar)>0 then
                v_Spell:=v_Spell||'Q';
            elsif instr(v_CharR,v_BitChar)>0 then
                v_Spell:=v_Spell||'R';
            elsif instr(v_CharS,v_BitChar)>0 then
                v_Spell:=v_Spell||'S';
            elsif instr(v_CharT,v_BitChar)>0 then
                v_Spell:=v_Spell||'T';
            elsif instr(v_CharW,v_BitChar)>0 then
                v_Spell:=v_Spell||'W';
            elsif instr(v_CharX,v_BitChar)>0 then
                v_Spell:=v_Spell||'X';
            elsif instr(v_CharY,v_BitChar)>0 then
                v_Spell:=v_Spell||'Y';
            elsif instr(v_CharZ,v_BitChar)>0 then
                v_Spell:=v_Spell||'Z';
            else
                v_Spell:=v_Spell||'_';
            end if;
            exit when length(v_Spell)>9;
        end loop;
    end if;
    RETURN(v_Spell);
END;
/
