CREATE OR REPLACE PROCEDURE ZL_���˽����¼_UPDATE(
	����ID_IN			����Ԥ����¼.����ID%TYPE,
	���ս���_IN			VARCHAR2,--"���㷽ʽ|������||....."
	����_IN				NUMBER:=0
) AS
	--���α�ΪҪɾ�����ɷ��ü�¼�����Ľ����¼
	CURSOR C_DEL IS
		SELECT * FROM ����Ԥ����¼ WHERE ����ID=����ID_IN;

	--���α������շѳ�Ԥ���Ŀ���Ԥ���б�(��SQL�ο�סԺ����)
	--��ID�������ȳ��ϴ�δ����ġ�
	CURSOR C_DEPOSIT(V_����ID ������Ϣ.����ID%TYPE) IS
	SELECT * FROM(
		SELECT A.ID,A.��¼״̬,A.NO,NVL(A.���,0) AS ���
		FROM ����Ԥ����¼ A,(
				SELECT NO,SUM(NVL(A.���,0)) AS ���
				FROM ����Ԥ����¼ A
			WHERE A.����ID IS NULL AND NVL(A.���,0)<>0 AND A.����ID=V_����ID
			  GROUP BY NO HAVING SUM(NVL(A.���,0))<>0
				) B
		WHERE A.����ID IS NULL AND NVL(A.���,0)<>0
		And A.���㷽ʽ Not IN (Select ���� From ���㷽ʽ Where ����=5)
		AND A.NO=B.NO AND A.����ID=V_����ID
		UNION ALL
		SELECT 0 AS ID,��¼״̬,NO,SUM(NVL(���,0)-NVL(��Ԥ��,0)) AS ���
		FROM ����Ԥ����¼
		WHERE ��¼���� IN(1,11) AND ����ID IS NOT NULL AND NVL(���,0)<>NVL(��Ԥ��,0) AND ����ID=V_����ID
		HAVING SUM(NVL(���,0)-NVL(��Ԥ��,0))<>0
		GROUP BY ��¼״̬,NO)
    ORDER BY ID,NO;

	--�����Ϣ
	V_NO			����Ԥ����¼.NO%TYPE;
	V_����ID		���˷��ü�¼.����ID%TYPE;
	V_��ҳID		���˷��ü�¼.��ҳID%TYPE;
	V_�Ǽ�ʱ��		���˷��ü�¼.�Ǽ�ʱ��%TYPE;
	V_����Ա���	���˷��ü�¼.����Ա���%TYPE;
	V_����Ա����	���˷��ü�¼.����Ա����%TYPE;

	--���ν������
	V_���ϼ�	����Ԥ����¼.��Ԥ��%TYPE;
	V_��Ԥ����	����Ԥ����¼.��Ԥ��%TYPE;

	V_��Ժ����	NUMBER;
	V_Ԥ�����  �������.Ԥ�����%TYPE;

	--���ս���
	V_���ս���	VARCHAR2(255);
	V_��ǰ����	VARCHAR2(50);
	V_���㷽ʽ	����Ԥ����¼.���㷽ʽ%TYPE;
	V_������	����Ԥ����¼.��Ԥ��%TYPE;
	V_�ֽ�		VARCHAR2(255);

	v_��¼����	����Ԥ����¼.��¼����%Type;

	--��ʱ����
	ERR_CUSTOM	EXCEPTION;
	V_ERROR		VARCHAR2(255);
BEGIN
	--ȡ�ñ��ν���������Ϣ
	IF NVL(����_IN,0)=1 THEN
		SELECT NO,����ID,�շ�ʱ��,����Ա���,����Ա����
			INTO V_NO,V_����ID,V_�Ǽ�ʱ��,V_����Ա���,V_����Ա����
		FROM ���˽��ʼ�¼ WHERE ID=����ID_IN;
	ELSE
		SELECT NO,����ID,�Ǽ�ʱ��,����Ա���,����Ա����
			INTO V_NO,V_����ID,V_�Ǽ�ʱ��,V_����Ա���,V_����Ա����
		FROM ���˷��ü�¼ WHERE ����ID=����ID_IN AND ROWNUM=1;

		Begin --20071027 �¶�
			Select ��¼���� Into v_��¼����
			From ����Ԥ����¼ Where ����ID=����ID_IN And Rownum=1;
		Exception --20071027 �¶�
			WHEN OTHERS Then v_��¼����:=-1; --20071027 �¶�
		End; --20071027 �¶�

	END IF;
	IF NVL(V_����ID,0)<>0 THEN
		SELECT סԺ���� INTO V_��ҳID FROM ������Ϣ WHERE ����ID=V_����ID;
	END IF;

	--�ж��Ƿ��Ժ����(Ԥ��ȫ������),�Ծ��������Ƿ�ȫ����Ԥ��
	V_��Ժ����:=0;
	IF ����_IN=1 THEN
		BEGIN
			SELECT Ԥ����� INTO V_Ԥ����� FROM ������� WHERE ����ID=V_����ID AND ����=1;
		EXCEPTION
			WHEN OTHERS THEN NULL;
		END;
		IF NVL(V_Ԥ�����,0)=0 THEN
			V_��Ժ����:=1;
		END IF ;
	END IF;

	--ɾ�����ν����ɷ��ó�������Ľ����¼
	V_���ϼ�:=0;V_��Ԥ����:=0;
	FOR R_DEL IN C_DEL LOOP
		IF R_DEL.��¼���� IN(1,11) THEN
			UPDATE �������
				SET Ԥ�����=NVL(Ԥ�����,0)+R_DEL.��Ԥ��
			WHERE ����ID=V_����ID AND ����=1;
			IF SQL%ROWCOUNT=0 THEN
				INSERT INTO �������(
					����ID,����,Ԥ�����,�������)
				VALUES(
					V_����ID,1,R_DEL.��Ԥ��,0);
			END IF;
			V_��Ԥ����:=V_��Ԥ����+R_DEL.��Ԥ��;
		ELSE
			UPDATE ��Ա�ɿ����
				SET ���=NVL(���,0)-R_DEL.��Ԥ��
			 WHERE �տ�Ա=V_����Ա���� AND ����=1
				AND ���㷽ʽ=R_DEL.���㷽ʽ;
			IF SQL%ROWCOUNT=0 THEN
				INSERT INTO ��Ա�ɿ����(
					�տ�Ա,���㷽ʽ,����,���)
				VALUES(
					V_����Ա����,R_DEL.���㷽ʽ,1,-1*R_DEL.��Ԥ��);
			END IF;
		END IF;

		V_���ϼ�:=V_���ϼ�+R_DEL.��Ԥ��;

		IF R_DEL.��¼����=1 THEN
			UPDATE ����Ԥ����¼ SET ��Ԥ��=NULL,����ID=NULL WHERE ID=R_DEL.ID;
		ELSE
			DELETE FROM ����Ԥ����¼ WHERE ID=R_DEL.ID;
		END IF;
	END LOOP;

	--------------------------------------------------------------------------------------------------------------
	--------------------------------------------------------------------------------------------------------------
	--����ҽ��֧������
	IF ���ս���_IN IS NOT NULL THEN
		--�������ս���
		V_���ս���:=���ս���_IN||'||';
		WHILE V_���ս��� IS NOT NULL LOOP
			V_��ǰ����:=SUBSTR(V_���ս���,1,INSTR(V_���ս���,'||')-1);

			V_���㷽ʽ:=SUBSTR(V_��ǰ����,1,INSTR(V_��ǰ����,'|')-1);
			V_������:=TO_NUMBER(SUBSTR(V_��ǰ����,INSTR(V_��ǰ����,'|')+1));

			INSERT INTO ����Ԥ����¼(
				ID,��¼����,NO,��¼״̬,����ID,��ҳID,ժҪ,���㷽ʽ,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
			VALUES(
				����Ԥ����¼_ID.NEXTVAL,DECODE(����_IN,1,2,v_��¼����),V_NO,1,V_����ID,V_��ҳID,'���ղ���',
				V_���㷽ʽ,V_�Ǽ�ʱ��,V_����Ա���,V_����Ա����,V_������,����ID_IN);

			V_���ϼ�:=V_���ϼ�-V_������;

			V_���ս���:=SUBSTR(V_���ս���,INSTR(V_���ս���,'||')+2);
		END LOOP;
	END IF;

	--���ʹ���˳�Ԥ��,���ȴ����Ԥ��(������)
	IF V_��Ԥ����<>0 THEN
		FOR R_DEPOSIT IN C_DEPOSIT(V_����ID) LOOP

			--���ʿ��Գ��Ľ��
			IF ����_IN=1 AND V_��Ժ����=1 THEN
				V_��Ԥ����:=R_DEPOSIT.���;
			ELSE
				IF R_DEPOSIT.���<V_���ϼ� THEN
					V_��Ԥ����:=R_DEPOSIT.���;
				ELSE
					V_��Ԥ����:=V_���ϼ�;
				END IF;
			END IF;

			IF R_DEPOSIT.ID<>0 THEN
				--��һ�γ�Ԥ��
				UPDATE ����Ԥ����¼
					SET ��Ԥ��=V_��Ԥ����,
						����ID=����ID_IN
				WHERE ID=R_DEPOSIT.ID;
			ELSE
				--���ϴ�ʣ���
				INSERT INTO ����Ԥ����¼(
					ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,���,
					���㷽ʽ,�������,ժҪ,�ɿλ,��λ������,��λ�ʺ�,�տ�ʱ��,
					����Ա����,����Ա���,��Ԥ��,����ID)
				SELECT ����Ԥ����¼_ID.NEXTVAL,NO,ʵ��Ʊ��,11,��¼״̬,����ID,
					 ��ҳID,����ID,NULL,���㷽ʽ,�������,ժҪ,�ɿλ,
					 ��λ������,��λ�ʺ�,�տ�ʱ��,����Ա����,����Ա���,
					 V_��Ԥ����,����ID_IN
				FROM ����Ԥ����¼
				WHERE NO=R_DEPOSIT.NO AND ��¼״̬=R_DEPOSIT.��¼״̬
					AND ��¼���� IN(1,11) AND ROWNUM=1;
			END IF;

			--����Ƿ��Ѿ�������
			V_���ϼ�:=V_���ϼ�-V_��Ԥ����;

			IF ����_IN=1 AND V_��Ժ����=1 THEN
				NULL;
			ELSE
				IF V_���ϼ�=0 THEN
					EXIT;
				END IF;
			END IF;
		END LOOP;
	END IF;

	--ʣ�ಿ��ȫ�����ֽ����
	IF V_���ϼ�<>0 THEN
		SELECT ���� INTO V_�ֽ� FROM ���㷽ʽ WHERE NVL(����,1)=1 AND ROWNUM<2;
		INSERT INTO ����Ԥ����¼(
			ID,��¼����,NO,��¼״̬,����ID,��ҳID,ժҪ,���㷽ʽ,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
		VALUES(
			����Ԥ����¼_ID.NEXTVAL,DECODE(����_IN,1,2,v_��¼����),V_NO,1,V_����ID,V_��ҳID,'�ֽ𲿷�',V_�ֽ�,
			V_�Ǽ�ʱ��,V_����Ա���,V_����Ա����,V_���ϼ�,����ID_IN);
	END IF;

	--����ٴ���"�������","��Ա�ɿ����"
	FOR R_DEL IN C_DEL LOOP
		IF R_DEL.��¼���� IN(1,11) THEN
			UPDATE �������
				SET Ԥ�����=NVL(Ԥ�����,0)-R_DEL.��Ԥ��
			WHERE ����ID=V_����ID AND ����=1;
			IF SQL%ROWCOUNT=0 THEN
				INSERT INTO �������(
					����ID,����,Ԥ�����,�������)
				VALUES(
					V_����ID,1,-1*R_DEL.��Ԥ��,0);
			END IF;
		ELSE
			UPDATE ��Ա�ɿ����
				SET ���=NVL(���,0)+R_DEL.��Ԥ��
			 WHERE �տ�Ա=V_����Ա���� AND ����=1
				AND ���㷽ʽ=R_DEL.���㷽ʽ;
			IF SQL%ROWCOUNT=0 THEN
				INSERT INTO ��Ա�ɿ����(
					�տ�Ա,���㷽ʽ,����,���)
				VALUES(
					V_����Ա����,R_DEL.���㷽ʽ,1,R_DEL.��Ԥ��);
			END IF;
		END IF;
	END LOOP;
	DELETE FROM ������� WHERE ����ID=V_����ID AND ����=1 AND NVL(�������,0)=0 AND NVL(Ԥ�����,0)=0;
	DELETE FROM ��Ա�ɿ���� WHERE ����=1 AND �տ�Ա=V_����Ա���� AND NVL(���,0)=0;
EXCEPTION
	WHEN ERR_CUSTOM THEN RAISE_APPLICATION_ERROR(-20101,'[ZLSOFT]'||V_ERROR||'[ZLSOFT]');
	WHEN OTHERS THEN ZL_ERRORCENTER(SQLCODE,SQLERRM);
END ZL_���˽����¼_UPDATE;
/

