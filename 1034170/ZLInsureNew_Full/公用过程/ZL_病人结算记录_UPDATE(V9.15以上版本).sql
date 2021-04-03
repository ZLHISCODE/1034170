CREATE OR REPLACE PROCEDURE ZL_病人结算记录_UPDATE(
	结帐ID_IN			病人预交记录.结帐ID%TYPE,
	保险结算_IN			VARCHAR2,--"结算方式|结算金额||....."
	结帐_IN				NUMBER:=0
) AS
	--该游标为要删除的由费用记录产生的结算记录
	CURSOR C_DEL IS
		SELECT * FROM 病人预交记录 WHERE 结帐ID=结帐ID_IN;

	--该游标用于收费冲预交的可用预交列表(该SQL参考住院结帐)
	--以ID排序，优先冲上次未冲完的。
	CURSOR C_DEPOSIT(V_病人ID 病人信息.病人ID%TYPE) IS
	SELECT * FROM(
		SELECT A.ID,A.记录状态,A.NO,NVL(A.金额,0) AS 金额
		FROM 病人预交记录 A,(
				SELECT NO,SUM(NVL(A.金额,0)) AS 金额
				FROM 病人预交记录 A
			WHERE A.结帐ID IS NULL AND NVL(A.金额,0)<>0 AND A.病人ID=V_病人ID
			  GROUP BY NO HAVING SUM(NVL(A.金额,0))<>0
				) B
		WHERE A.结帐ID IS NULL AND NVL(A.金额,0)<>0
		And A.结算方式 Not IN (Select 名称 From 结算方式 Where 性质=5)
		AND A.NO=B.NO AND A.病人ID=V_病人ID
		UNION ALL
		SELECT 0 AS ID,记录状态,NO,SUM(NVL(金额,0)-NVL(冲预交,0)) AS 金额
		FROM 病人预交记录
		WHERE 记录性质 IN(1,11) AND 结帐ID IS NOT NULL AND NVL(金额,0)<>NVL(冲预交,0) AND 病人ID=V_病人ID
		HAVING SUM(NVL(金额,0)-NVL(冲预交,0))<>0
		GROUP BY 记录状态,NO)
    ORDER BY ID,NO;

	--相关信息
	V_NO			病人预交记录.NO%TYPE;
	V_病人ID		病人费用记录.病人ID%TYPE;
	V_主页ID		病人费用记录.主页ID%TYPE;
	V_登记时间		病人费用记录.登记时间%TYPE;
	V_操作员编号	病人费用记录.操作员编号%TYPE;
	V_操作员姓名	病人费用记录.操作员姓名%TYPE;

	--本次结算变量
	V_金额合计	病人预交记录.冲预交%TYPE;
	V_冲预交额	病人预交记录.冲预交%TYPE;

	V_出院结帐	NUMBER;
	V_预交余额  病人余额.预交余额%TYPE;

	--保险结算
	V_保险结算	VARCHAR2(255);
	V_当前结算	VARCHAR2(50);
	V_结算方式	病人预交记录.结算方式%TYPE;
	V_结算金额	病人预交记录.冲预交%TYPE;
	V_现金		VARCHAR2(255);

	v_记录性质	病人预交记录.记录性质%Type;

	--临时变量
	ERR_CUSTOM	EXCEPTION;
	V_ERROR		VARCHAR2(255);
BEGIN
	--取得本次结算的相关信息
	IF NVL(结帐_IN,0)=1 THEN
		SELECT NO,病人ID,收费时间,操作员编号,操作员姓名
			INTO V_NO,V_病人ID,V_登记时间,V_操作员编号,V_操作员姓名
		FROM 病人结帐记录 WHERE ID=结帐ID_IN;
	ELSE
		SELECT NO,病人ID,登记时间,操作员编号,操作员姓名
			INTO V_NO,V_病人ID,V_登记时间,V_操作员编号,V_操作员姓名
		FROM 病人费用记录 WHERE 结帐ID=结帐ID_IN AND ROWNUM=1;

		Begin --20071027 陈东
			Select 记录性质 Into v_记录性质
			From 病人预交记录 Where 结帐ID=结帐ID_IN And Rownum=1;
		Exception --20071027 陈东
			WHEN OTHERS Then v_记录性质:=-1; --20071027 陈东
		End; --20071027 陈东

	END IF;
	IF NVL(V_病人ID,0)<>0 THEN
		SELECT 住院次数 INTO V_主页ID FROM 病人信息 WHERE 病人ID=V_病人ID;
	END IF;

	--判断是否出院结帐(预交全部冲完),以决定后面是否全部冲预交
	V_出院结帐:=0;
	IF 结帐_IN=1 THEN
		BEGIN
			SELECT 预交余额 INTO V_预交余额 FROM 病人余额 WHERE 病人ID=V_病人ID AND 性质=1;
		EXCEPTION
			WHEN OTHERS THEN NULL;
		END;
		IF NVL(V_预交余额,0)=0 THEN
			V_出院结帐:=1;
		END IF ;
	END IF;

	--删除本次结帐由费用程序产生的结算记录
	V_金额合计:=0;V_冲预交额:=0;
	FOR R_DEL IN C_DEL LOOP
		IF R_DEL.记录性质 IN(1,11) THEN
			UPDATE 病人余额
				SET 预交余额=NVL(预交余额,0)+R_DEL.冲预交
			WHERE 病人ID=V_病人ID AND 性质=1;
			IF SQL%ROWCOUNT=0 THEN
				INSERT INTO 病人余额(
					病人ID,性质,预交余额,费用余额)
				VALUES(
					V_病人ID,1,R_DEL.冲预交,0);
			END IF;
			V_冲预交额:=V_冲预交额+R_DEL.冲预交;
		ELSE
			UPDATE 人员缴款余额
				SET 余额=NVL(余额,0)-R_DEL.冲预交
			 WHERE 收款员=V_操作员姓名 AND 性质=1
				AND 结算方式=R_DEL.结算方式;
			IF SQL%ROWCOUNT=0 THEN
				INSERT INTO 人员缴款余额(
					收款员,结算方式,性质,余额)
				VALUES(
					V_操作员姓名,R_DEL.结算方式,1,-1*R_DEL.冲预交);
			END IF;
		END IF;

		V_金额合计:=V_金额合计+R_DEL.冲预交;

		IF R_DEL.记录性质=1 THEN
			UPDATE 病人预交记录 SET 冲预交=NULL,结帐ID=NULL WHERE ID=R_DEL.ID;
		ELSE
			DELETE FROM 病人预交记录 WHERE ID=R_DEL.ID;
		END IF;
	END LOOP;

	--------------------------------------------------------------------------------------------------------------
	--------------------------------------------------------------------------------------------------------------
	--产生医保支付结算
	IF 保险结算_IN IS NOT NULL THEN
		--各个保险结算
		V_保险结算:=保险结算_IN||'||';
		WHILE V_保险结算 IS NOT NULL LOOP
			V_当前结算:=SUBSTR(V_保险结算,1,INSTR(V_保险结算,'||')-1);

			V_结算方式:=SUBSTR(V_当前结算,1,INSTR(V_当前结算,'|')-1);
			V_结算金额:=TO_NUMBER(SUBSTR(V_当前结算,INSTR(V_当前结算,'|')+1));

			INSERT INTO 病人预交记录(
				ID,记录性质,NO,记录状态,病人ID,主页ID,摘要,结算方式,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
			VALUES(
				病人预交记录_ID.NEXTVAL,DECODE(结帐_IN,1,2,v_记录性质),V_NO,1,V_病人ID,V_主页ID,'保险部份',
				V_结算方式,V_登记时间,V_操作员编号,V_操作员姓名,V_结算金额,结帐ID_IN);

			V_金额合计:=V_金额合计-V_结算金额;

			V_保险结算:=SUBSTR(V_保险结算,INSTR(V_保险结算,'||')+2);
		END LOOP;
	END IF;

	--如果使用了冲预交,则先处理冲预交(尽量冲)
	IF V_冲预交额<>0 THEN
		FOR R_DEPOSIT IN C_DEPOSIT(V_病人ID) LOOP

			--本笔可以冲款的金额
			IF 结帐_IN=1 AND V_出院结帐=1 THEN
				V_冲预交额:=R_DEPOSIT.金额;
			ELSE
				IF R_DEPOSIT.金额<V_金额合计 THEN
					V_冲预交额:=R_DEPOSIT.金额;
				ELSE
					V_冲预交额:=V_金额合计;
				END IF;
			END IF;

			IF R_DEPOSIT.ID<>0 THEN
				--第一次冲预交
				UPDATE 病人预交记录
					SET 冲预交=V_冲预交额,
						结帐ID=结帐ID_IN
				WHERE ID=R_DEPOSIT.ID;
			ELSE
				--冲上次剩余额
				INSERT INTO 病人预交记录(
					ID,NO,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,金额,
					结算方式,结算号码,摘要,缴款单位,单位开户行,单位帐号,收款时间,
					操作员姓名,操作员编号,冲预交,结帐ID)
				SELECT 病人预交记录_ID.NEXTVAL,NO,实际票号,11,记录状态,病人ID,
					 主页ID,科室ID,NULL,结算方式,结算号码,摘要,缴款单位,
					 单位开户行,单位帐号,收款时间,操作员姓名,操作员编号,
					 V_冲预交额,结帐ID_IN
				FROM 病人预交记录
				WHERE NO=R_DEPOSIT.NO AND 记录状态=R_DEPOSIT.记录状态
					AND 记录性质 IN(1,11) AND ROWNUM=1;
			END IF;

			--检查是否已经处理完
			V_金额合计:=V_金额合计-V_冲预交额;

			IF 结帐_IN=1 AND V_出院结帐=1 THEN
				NULL;
			ELSE
				IF V_金额合计=0 THEN
					EXIT;
				END IF;
			END IF;
		END LOOP;
	END IF;

	--剩余部份全部用现金结算
	IF V_金额合计<>0 THEN
		SELECT 名称 INTO V_现金 FROM 结算方式 WHERE NVL(性质,1)=1 AND ROWNUM<2;
		INSERT INTO 病人预交记录(
			ID,记录性质,NO,记录状态,病人ID,主页ID,摘要,结算方式,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
		VALUES(
			病人预交记录_ID.NEXTVAL,DECODE(结帐_IN,1,2,v_记录性质),V_NO,1,V_病人ID,V_主页ID,'现金部份',V_现金,
			V_登记时间,V_操作员编号,V_操作员姓名,V_金额合计,结帐ID_IN);
	END IF;

	--最后再处理"病人余额","人员缴款余额"
	FOR R_DEL IN C_DEL LOOP
		IF R_DEL.记录性质 IN(1,11) THEN
			UPDATE 病人余额
				SET 预交余额=NVL(预交余额,0)-R_DEL.冲预交
			WHERE 病人ID=V_病人ID AND 性质=1;
			IF SQL%ROWCOUNT=0 THEN
				INSERT INTO 病人余额(
					病人ID,性质,预交余额,费用余额)
				VALUES(
					V_病人ID,1,-1*R_DEL.冲预交,0);
			END IF;
		ELSE
			UPDATE 人员缴款余额
				SET 余额=NVL(余额,0)+R_DEL.冲预交
			 WHERE 收款员=V_操作员姓名 AND 性质=1
				AND 结算方式=R_DEL.结算方式;
			IF SQL%ROWCOUNT=0 THEN
				INSERT INTO 人员缴款余额(
					收款员,结算方式,性质,余额)
				VALUES(
					V_操作员姓名,R_DEL.结算方式,1,R_DEL.冲预交);
			END IF;
		END IF;
	END LOOP;
	DELETE FROM 病人余额 WHERE 病人ID=V_病人ID AND 性质=1 AND NVL(费用余额,0)=0 AND NVL(预交余额,0)=0;
	DELETE FROM 人员缴款余额 WHERE 性质=1 AND 收款员=V_操作员姓名 AND NVL(余额,0)=0;
EXCEPTION
	WHEN ERR_CUSTOM THEN RAISE_APPLICATION_ERROR(-20101,'[ZLSOFT]'||V_ERROR||'[ZLSOFT]');
	WHEN OTHERS THEN ZL_ERRORCENTER(SQLCODE,SQLERRM);
END ZL_病人结算记录_UPDATE;
/

