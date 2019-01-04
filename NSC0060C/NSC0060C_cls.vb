Option Strict Off
Option Explicit On
Imports General
Imports system.data
<System.Runtime.InteropServices.ProgId("clsStaffIdo_NET.clsStaffIdo")> Public Class clsStaffIdo
	'/****************************************************/
	'/    ���і��́F�i�[�X�G�C�h6.0
	'/ ��۸��і��́F�E���f�[�^�擾���i
	'/        �h�c�FNsc0000
	'/        �T�v�F�E���̏����擾����B
	'/
	'/      �쐬�ҁF D.T     CREATE 2006/02/27    REV 01.00
	'/      �X�V�ҁF M.J     UPDATE 2009/06/09    REV 01.01
	'/                         �X�V���e�F(PKG-0122)
	'/      �X�V�ҁF M.J     UPDATE 2009/06/13    REV 01.02
	'/                         �X�V���e�F(PKG-0207)
	'/      �X�V�ҁF T.I     UPDATE 2009/08/03    REV 01.03
	'/                         �X�V���e�F(P-02074)
    '/      �X�V�ҁF okamura        2009/08/19   �yP-02212�z
    '/      �X�V�ҁF T.Sasaki       2012/02/13   �yP-04934�z
    '/      �X�V�ҁF Fujisawa       2012/10/25   �yP-05558�zver7.0�Ή�
    '/      �X�V�ҁF murakami       2012/11/29   �yP-05639�z
    '/      �쐬�ҁF M.I            2012/12/17   �yP-05377�z
    '/      �X�V�ҁF T.K            2018/08/24   �yP-09479�z
    '/
	'/     Copyright (C) Inter co.,ltd 2002
	'/****************************************************/

#Region "�ϐ��錾"


    Public m_strHospitalCD As String = String.Empty '�{�݃R�[�h
    Public m_strStaffMngID As String = String.Empty '�E���Ǘ��ԍ�
    Public m_numDateFlg As Integer '���t�t���O(0:�P����A1:���Ԏw��)
    Public m_numDateFrom As Integer '�J�n�N����(YYYYMMDD)
    Public m_numDateTo As Integer '�I���N����(YYYYMMDD)
    Public m_numSortFlg As Integer '�\�[�g�t���O(0:�����A1:�~��)
    Public m_numNendo As Integer '�N�x
    Public m_intDelKbn As Short '�폜�󋵁i0:���폜,1:�폜,2:�S�āj
    Public m_numSaiyoKensu As Integer '�̗p�ٓ���������
    Public m_numSaiyoIdx As Integer '�̗p�ٓ������C���f�b�N�X
    Public m_numKinmuDeptKensu As Integer '�Ζ������ٓ�����
    Public m_numKinmuDeptIdx As Integer '�Ζ������ٓ��C���f�b�N�X
    Public m_numWardDeptKensu As Integer '�z�������ٓ�����
    Public m_numWardDeptIdx As Integer '�z�������ٓ��C���f�b�N�X
    Public m_numPostKensu As Integer '��E�ٓ�����
    Public m_numPostIdx As Integer '��E�ٓ��C���f�b�N�X
    Public m_numJobKensu As Integer '�E��ٓ�����
    Public m_numJobIdx As Integer '�E��ٓ��C���f�b�N�X
    Public m_numKenmuKensu As Integer '�����ٓ�����
    Public m_numKenmuIdx As Integer '�����ٓ��C���f�b�N�X
    Public m_numSaikeiKensu As Integer '�Čf�ٓ�����
    Public m_numSaikeiIdx As Integer '�Čf�ٓ��C���f�b�N�X
    Public m_numMenkyoKensu As Integer '�Ƌ���񌏐�
    Public m_numMenkyoIdx As Integer '�Ƌ����C���f�b�N�X
    Public m_numShikakuKensu As Integer '���i��񌏐�
    Public m_numShikakuIdx As Integer '���i���C���f�b�N�X
    Public m_numIinKensu As Integer '�ψ���񌏐�
    Public m_numIinIdx As Integer '�ψ����C���f�b�N�X
    Public m_numSyokurekiKensu As Integer '�E����񌏐�
    Public m_numSyokurekiIdx As Integer '�E�����C���f�b�N�X
    Public m_numIppanGakurekiKensu As Integer '��ʊw����񌏐�
    Public m_numIppanGakurekiIdx As Integer '��ʊw�����C���f�b�N�X
    Public m_numSenmonGakurekiKensu As Integer '���w����񌏐�
    Public m_numSenmonGakurekiIdx As Integer '���w�����C���f�b�N�X
    Public m_numChoukyuKensu As Integer '���x��񌏐�
    Public m_numChoukyuIdx As Integer '���x���C���f�b�N�X
    Public m_numSankyuKensu As Integer '�Y�x��񌏐�
    Public m_numSankyuIdx As Integer '�Y�x���C���f�b�N�X
    Public m_numKyoukaiKensu As Integer '�����񌏐�
    Public m_numKyoukaiIdx As Integer '������C���f�b�N�X
    Public m_numKazokuKensu As Integer '�Ƒ���񌏐�
    Public m_numKazokuIdx As Integer '�Ƒ����C���f�b�N�X
    Public m_numStudyKensu As Integer '���C��񌏐�
    Public m_numStudyIdx As Integer '���C���C���f�b�N�X
    Public m_numStudyDateKensu As Integer '���C���t��񌏐�
    Public m_numStudyDateIdx As Integer '���C���t���C���f�b�N�X
    Public m_numGyosekiKensu As Integer '�Ɛь���
    Public m_numGyosekiIdx As Integer '�ƐуC���f�b�N�X
    '2012/02/13 Sasaki add start---------------------------------------------------------------
    Public m_lngSAIdx As Long '���C��u���C���f�b�N�X
    Public m_lngSACount As Long '���C��u������
    Public m_numShortTimeKensu As Integer '�Z���Ԑ��x��񌏐�
    Public m_numShortTimeIdx As Integer '�Z���Ԑ��x���C���f�b�N�X
    Public m_numNightWorkerKensu As Integer '��ΐ�]��񌏐�
    Public m_numNightWorkerIdx As Integer '��ΐ�]���C���f�b�N�X
    Public m_numHealthCondHisIdx As Integer '���N��ԗ������C���f�b�N�X
    Public m_numHBChkHisInfoIdx As Integer '���N��ԗ������C���f�b�N�X
    Public m_numKansensyouHisIdx As Integer '�����Ǘ��C���f�b�N�X

    Public Structure SA_Type
        Dim lngNendo As Long '�N�x
        Dim lngStudyIndex As Long '���C�h�m�c�d�w
        Dim strOutInFlg As String '�@���O�t���O
        Dim strStudyCD As String '���C�R�[�h
        Dim strStudyName As String '���C����
        Dim strStudySecName As String '���C����
        Dim strStudyKana As String '���C�t���K�i
        Dim strKindCD As String '��ރR�[�h
        Dim strKindName As String '��ޖ���
        Dim strSponsorCD As String '��ÃR�[�h
        Dim strSponsorName As String '��Ö���
        Dim strTheme As String '�e�[�}
        Dim strLecturer As String '�u�t
        Dim strHall As String '���E�ꏊ
        Dim strSankaCond As String '�Q������
        Dim lngSankaNinzu As Long '�Q���l��
        Dim strReports As String '�A������
        Dim strBikou As String '���l
        Dim strUrl As String '�t�q�k
        Dim strNecessaryValuationLevelCD As String '�K�{�]�����x���R�[�h
        Dim strNecessaryValuationLevelName As String '�K�{�]�����x������
        Dim strNecessaryValuationLevelSecName As String '�K�{�]�����x������
        Dim strNecessaryValuationLevelMark As String '�K�{�]�����x���L��
        Dim lngAcceptFromDate As Long '��t�J�n�N����
        Dim lngAcceptToDate As Long '��t�I���N����
        Dim strAcceptapstate As String '��t�\���݃t���O
        Dim strNendoPlanKbn As String '�N�Ԍv��敪
        Dim strKinmuDeptCD As String '�Ζ������R�[�h
        Dim strKinmuDeptName As String '�Ζ���������
        Dim strAllDaysNecessaryFlg As String '�S�����K�{�t���O
        Dim strIndependentFlg As String '���匤�C�t���O
        Dim lngDateIdx As Long '���t�C���f�b�N�X
        Dim strDateAppoFlg As String '���t�w��t���O
        Dim lngDateFrom As Long '���t�J�n�N����
        Dim lngDateTo As Long '���t�I���N����
        Dim strJapanAreaCD As String '�s���{���R�[�h
        Dim strAttendCompFlg As String '��u�σt���O
        Dim strAttendLecrep As String '��u��
        Dim strCostCD As String '��p�R�[�h
        Dim strCostName As String '��p����
        Dim strSankaFormCD As String '�Q���`�ԃR�[�h
        Dim strSankaFormName As String '�Q���`�Ԗ���
        Dim strSSBikou As String '���l
        Dim strUniqueSeqNo As String 'UNIQUESEQNO
        Dim strApproveFlg As String '���F�σt���O
        Dim strSankaFlg As String '�Q���t���O
        Dim strJapanAreaName As String '�s���{������
        Dim dblRegistFirstTimeDate As Double '����o�^����
        Dim dblLastUpdTimeDate As Double '�ŏI�X�V����
        Dim strRegistrantID As String '�o�^�҂h�c
    End Structure
    '2012/02/13 Sasaki add end-----------------------------------------------------------------

    '�ٓ���
    Public Structure Saiyo_Type
        Dim strEmpCD As String '�̗p�R�[�h
        Dim strEmpName As String '�̗p����
        Dim strEmpSecName As String '�̗p����
        Dim numEmpDate As Integer '�̗p�N����
        Dim strRetireCD As String '�ސE�R�[�h
        Dim strRetireName As String '�ސE����
        Dim numRetireDate As Integer '�ސE�N����
        Dim lngFirstTime As Long '����o�^����
        Dim lngUpdTime As Long '�ŏI�X�V����
        Dim strStaffID As String '�E���ԍ�   
    End Structure
    Public Structure Ido_Type
        Dim strCD As String '�R�[�h
        Dim strName As String '����
        Dim numDateFrom As Integer '�J�n�N����
        Dim numDateTo As Integer '�I���N����
        Dim lngFirstTime As Long '����o�^����
        Dim lngUpdTime As Long '�ŏI�X�V����
        Dim strIdoHope As String '�ٓ���]�t���O(1:�ٓ���]���� 0:�ٓ���]����)
        Dim SecName As String
        Dim DispNo As Integer
    End Structure
    Public Structure Kenmu_Type
        Dim strWardDeptCD As String '�z�������R�[�h
        Dim strWardDeptName As String '�z����������
        Dim strKinmuDeptCD As String '�Ζ������R�[�h
        Dim strKinmuDeptName As String '�Ζ���������
        Dim strPostCD As String '��E�R�[�h
        Dim strPostName As String '��E����
        Dim numDateFrom As Integer '�J�n�N����
        Dim numSEQ As Integer 'SEQ
        Dim numDateTo As Integer '�I���N����
        Dim lngFirstTime As Long '����o�^����
        Dim lngUpdTime As Long '�ŏI�X�V����
    End Structure

    '�̗p�����^�C�v
    Public Structure SaiyoIdo_Type
        Dim strHospitalCD As String '�{�݃R�[�h
        Dim strStaffMngID As String '�E���Ǘ��ԍ�
        Dim Ido() As Saiyo_Type '�̗p�ٓ����
    End Structure

    Public g_SaiyoIdo As SaiyoIdo_Type

    Public Structure StaffIdo_Type
        Dim strHospitalCD As String '�{�݃R�[�h
        Dim strStaffMngID As String '�E���Ǘ��ԍ�
        Dim Ido() As Ido_Type '�ٓ����
    End Structure

    '�����ٓ��^�C�v
    Public Structure KenmuIdo_Type
        Dim strHospitalCD As String '�{�݃R�[�h
        Dim strStaffMngID As String '�E���Ǘ��ԍ�
        Dim Ido() As Kenmu_Type '�ٓ����
    End Structure

    '�Ƌ��E���i���^�C�v
    Public Structure MenkyoInfo_Type
        Dim strHospitalCD As String '�{�݃R�[�h
        Dim strStaffMngID As String '�E���Ǘ��ԍ�
        Dim strCD As String '�R�[�h
        Dim strName As String '����
        Dim strSecName As String '2018/10/02 Darren ADD
        Dim strNo As String '�ԍ�

        '2012/10/25 fujisawa add st ----------------
        Dim strJapanAreaCD As String '�s���{���R�[�h�@
        Dim strJapanAreaName As String '�s���{������
        '2012/10/25 fujisawa add end ---------------

        '�ψ��p
        Dim strIinPostCd As String '��ECD
        Dim strIinPostName As String '��E��

        '2018/08/24 T.K add st ---------------------
        '���x�p
        Dim numWeeklyTime As Integer '�T�J������
        '2018/08/24 T.K add ed ---------------------

        Dim numGetDate As Integer '�擾��
        Dim numEndDate As Integer '�މ�N����
        Dim numDateFrom As Integer '�J�n�N����
        Dim numDateTo As Integer '�I���N����
        Dim strBikou As String '���l
        Dim lngFirstTime As Long '����o�^����
        Dim lngUpdTime As Long '�ŏI�X�V����
    End Structure

    '�E���^�C�v
    Public Structure SyokurekiInfo_Type
        Dim strHospitalCD As String '�{�݃R�[�h
        Dim strStaffMngID As String '�E���Ǘ��ԍ�
        Dim strCD As String '�R�[�h
        Dim strName As String '����
        Dim strArea As String '�Ζ��@��
        Dim numGetDate As Integer '�擾��
        Dim numDateFrom As Integer '�J�n�N����
        Dim numDateTo As Integer '�I���N����
        Dim strExpMedicalName As String '�o���f�É� 
        Dim strBikou As String '���l
        Dim lngFirstTime As Long '����o�^����
        Dim lngUpdTime As Long '�ŏI�X�V����
    End Structure

    '�w���^�C�v
    Public Structure GakurekiInfo_Type
        Dim strHospitalCD As String '�{�݃R�[�h
        Dim strStaffMngID As String '�E���Ǘ��ԍ�
        Dim strKbn As String '�敪
        Dim strKbnName As String '�敪����
        Dim strChiikiCD As String '�n��R�[�h
        Dim strChiikiName As String '�n�於��
        Dim strLastKbn As String '�ŏI�w���敪
        Dim numDate As Integer '���ƔN����
        Dim strSchoolCD As String '�w�Z�R�[�h
        Dim strSchoolName As String '�w�Z��
        Dim strBikou As String '�C���ߒ�
        Dim lngFirstTime As Long '����o�^����
        Dim lngUpdTime As Long '�ŏI�X�V����
    End Structure

    Public Structure SankyuInfo_Type
        Dim strHospitalCD As String '�{�݃R�[�h
        Dim strStaffMngID As String '�E���Ǘ��ԍ�
        Dim numSEQ As Integer 'SEQ
        Dim numPlanDate As Integer '�\��N����
        Dim strTwinFlg As String '�o�ً敪
        Dim numBirthDate As Integer '�o�Y�N����
        Dim numPlanSanzenYamenFrom As Integer '�\��Y�O���From
        Dim numPlanSanzenYamenTo As Integer '�\��Y�O���To
        Dim numPlanSanzenHolFrom As Integer '�\��Y�O�x��From
        Dim numPlanSanzenHolTo As Integer '�\��Y�O�x��To
        Dim numPlanSangoHolFrom As Integer '�\��Y��x��From
        Dim numPlanSangoHolTo As Integer '�\��Y��x��To
        Dim numPlanIkujiHolFrom As Integer '�\��玙�x��From
        Dim numPlanIkujiHolTo As Integer '�\��玙�x��To
        Dim numFixedSanzenYamenFrom As Integer '�m��Y�O���From
        Dim numFixedSanzenYamenTo As Integer '�m��Y�O���To
        Dim numFixedSanzenHolFrom As Integer '�m��Y�O�x��From
        Dim numFixedSanzenHolTo As Integer '�m��Y�O�x��To
        Dim numFixedSangoHolFrom As Integer '�m��Y��x��From
        Dim numFixedSangoHolTo As Integer '�m��Y��x��To
        Dim numFixedIkujiHolFrom As Integer '�m��玙�x��From
        Dim numFixedIkujiHolTo As Integer '�m��玙�x��To
        Dim strUniqueSeqNO As String 'UNIQUESEQNO
        Dim strApproveFlg As String '���F�σt���O
        Dim lngFirstTime As Long '����o�^����
        Dim lngUpdTime As Long '�ŏI�X�V����
    End Structure

    Public Structure Kazoku_Type
        Dim strHospitalCD As String '�{�݃R�[�h
        Dim strStaffMngID As String '�E���Ǘ��ԍ�
        Dim numSEQ As Integer 'SEQ
        Dim strName As String '����
        Dim numDate As Integer '���N����
        Dim strTsudukiCD As String '�����R�[�h
        Dim strTsudukiName As String '������
        Dim strDoukyoKbn As String '�����敪
        Dim strFuyouKbn As String '�}�{�敪
        Dim strSeizonKbn As String '�����敪
        Dim lngFirstTime As Long '����o�^����
        Dim lngUpdTime As Long '�ŏI�X�V����
    End Structure


    Public Structure StudySub_Type
        Dim numFromDate As Integer '�J�n��
        Dim numToDate As Integer '�I����
        Dim strDateType As String '���Ԃ̃^�C�v
    End Structure

    '���C���e�@�\����
    Public Structure Study_Type
        Dim numYYYY As Integer '�N�x
        Dim strSEQ As String 'SEQ No.
        Dim strCourseCD As String '�������
        Dim strKbnCD As String '�敪����
        Dim strSyuruiCD As String '��޺���
        Dim strSyusaiCD As String '��ú���
        Dim strSankaCD As String '�Q������
        Dim strApplyStatus As String '��u��
        Dim strDeleteStatus As String '�폜��
        Dim strApplyRepo As String '��u��
        Dim strBiko As String '���l
        Dim numNewFlg As Integer '����/�V�K�ް� �����׸�
        Dim numProcFlg As Integer '���� �����׸�
        Dim numDispIndex As Integer 'ؽĕ\�����̲��ޯ��
        'Ͻ�����擾�����e����
        Dim strCorseName As String '�R�[�X�@����
        Dim strKbnName As String '���C�敪�@����
        Dim strSyuruiName As String '��ށ@����
        Dim strSyusaiName As String '��Á@����
        Dim strSankaName As String '�Q��   ����
        Dim strThema As String '�e�[�}
        Dim strPlaningFLG As String '�v��FLG
        Dim strCostCD As String '��pCD
        Dim strCostName As String '��p����
        Dim strCostCD2 As String '��pCD(���C�\��F)
        Dim strCostName2 As String '��p����(���C�\��F)
        Dim strDate As String '��������܂Ƃ߂ɂ�������
        Dim objDateList() As StudySub_Type '�e���C�ɑ΂������
    End Structure

    '�Z���ԁ���ΐ�]�@�\����
    Public Structure Worker_Type
        Dim hospCd As String '�a�@�R�[�h
        Dim mngId As String '�E���Ǘ��ԍ�
        Dim dateFrom As Integer '�J�n��
        Dim dateTo As Integer '�I����
        Dim reasonCd As String '���R�R�[�h
        Dim name As String '����
        Dim secName As String '����
        Dim birthDate As Integer '�o�Y��
        Dim fstRegDate As Long '����o�^����
        Dim lstUpdDate As Long '�ŏI�X�V����
        Dim lstUserId As String '�ŏI�o�^��
    End Structure

    Public g_KinmuDeptIdo As StaffIdo_Type '�Ζ������ٓ�
    Public g_WardDeptIdo As StaffIdo_Type '�z�������ٓ�
    Public g_PostIdo As StaffIdo_Type '��E�ٓ�
    Public g_JobIdo As StaffIdo_Type '�E��ٓ�
    Public g_KenmuIdo As KenmuIdo_Type '�����ٓ�
    Public g_MenkyoInfo() As MenkyoInfo_Type '�Ƌ����
    Public g_ShikakuInfo() As MenkyoInfo_Type '���i���
    Public g_IinInfo() As MenkyoInfo_Type '�ψ����
    Public g_SyokurekiInfo() As SyokurekiInfo_Type '�E�����
    Public g_IppanGakurekiInfo() As GakurekiInfo_Type '��ʊw�����
    Public g_SenmonGakurekiInfo() As GakurekiInfo_Type '���w�����
    Public g_ChoukyuInfo() As MenkyoInfo_Type '���x���
    Public g_SankyuInfo() As SankyuInfo_Type '�Y�x���
    Public g_KyoukaiInfo() As MenkyoInfo_Type '������
    Public g_KazokuInfo() As Kazoku_Type '�Ƒ����
    Public g_SaikeiIdo As StaffIdo_Type '�Čf�ٓ�   
    Public g_StudyInfo() As Study_Type
    Public g_ShortTimeIdo() As Worker_Type '�Z����
    Public g_NightWorkerIdo() As Worker_Type '��ΐ�]
    '2012/02/13 Sasaki add start--------------------------------
    Public g_StudyAttend() As SA_Type '���C��u
    Public m_strSankaFlg As String '�Q���t���O
    Public m_strAttendCompFlg As String '��u�σt���O
    Public m_strApproveFlg As String '���F�σt���O
    Public m_strOutInFlg As String '�@���O�t���O
    '2012/02/13 Sasaki add end----------------------------------

    '�Ɛя��^�C�v
    Public Structure Gyoseki_Type
        Dim strHospitalCD As String '�{�݃R�[�h
        Dim strStaffMngID As String '�E���Ǘ��ԍ�
        Dim strGyosekiCd As String '�ƐуR�[�h
        Dim strGyosekiName As String '�Ɛі���
        Dim numSEQ As Integer 'SEQ
        Dim numFromDate As Integer '�J�n�N����
        Dim numToDate As String '�I���N����
        Dim strSubject As String '����
        Dim strGyosekiPlaceCd As String '�Ɛє��\�ꏊ�R�[�h
        Dim strGyosekiPlaceName As String '�Ɛє��\�ꏊ����
        Dim strGyosekiBikou As String '�Ɛє��l
        Dim dblRegistFirstTimeDate As Double '����o�^����
        Dim dblLastUpdTimeDate As Double '�ŏI�X�V����(�r���p�j
    End Structure
    Public g_Gyoseki() As Gyoseki_Type

    '���N��ԗ���
    Private g_HealthCondHis As DataTable
    '�g�a�����������
    Private g_HBChkHisInfo As DataTable
    '�����Ǘ�
    Private g_KansensyouHis As DataTable

#End Region

    ''' <summary>
    ''' �̗p�������擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetSaiyoIdo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetSaiyoIdo"

            mGetSaiyoIdo = False

            '�擾
            If fncGetSaiyoIdo() = False Then
                Exit Function
            End If

            '����I��
            mGetSaiyoIdo = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ζ������ٓ������擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetKinmuDeptIdo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetKinmuDeptIdo"

            mGetKinmuDeptIdo = False

            '�擾
            If fncGetKinmuDeptIdo() = False Then
                Exit Function
            End If

            '����I��
            mGetKinmuDeptIdo = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Čf�ٓ������擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetSaikeiIdo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetSaikeiIdo"

            mGetSaikeiIdo = False

            '�擾
            If fncGetSaikeiIdo() = False Then
                Exit Function
            End If

            '����I��
            mGetSaikeiIdo = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �z�������ٓ������擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetWardDeptIdo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetWardDeptIdo"
            mGetWardDeptIdo = False
            '�擾
            If fncGetWardDeptIdo() = False Then
                Exit Function
            End If

            '����I��
            mGetWardDeptIdo = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��E�ٓ������擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetPostIdo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetPostIdo"

            mGetPostIdo = False

            '�擾
            If fncGetPostIdo() = False Then
                Exit Function
            End If

            '����I��
            mGetPostIdo = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E��ٓ������擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetJobIdo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetJobIdo"

            mGetJobIdo = False

            '�擾
            If fncGetJobIdo() = False Then
                Exit Function
            End If

            '����I��
            mGetJobIdo = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���������擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetKenmuIdo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetKenmuIdo"

            mGetKenmuIdo = False

            '�擾
            If fncGetKenmuIdo() = False Then
                Exit Function
            End If

            '����I��
            mGetKenmuIdo = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ƌ������擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetMenkyoInfo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetMenkyoInfo"

            mGetMenkyoInfo = False

            '�擾
            If fncGetMenkyoInfo() = False Then
                Exit Function
            End If

            '����I��
            mGetMenkyoInfo = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���i�����擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetShikakuInfo() As Boolean

        General.g_ErrorProc = "clsStaffIdo mGetShikakuInfo"

        mGetShikakuInfo = False
        Try
            '�擾
            If fncGetShikakuInfo() = False Then
                Exit Function
            End If

            '����I��
            mGetShikakuInfo = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ψ������擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetIinInfo() As Boolean

        General.g_ErrorProc = "clsStaffIdo mGetIinInfo"

        mGetIinInfo = False
        Try
            '�擾
            If fncGetIinInfo() = False Then
                Exit Function
            End If

            '����I��
            mGetIinInfo = True

        Catch ex As Exception
            Throw
        End Try

    End Function

    ''' <summary>
    ''' �E�������擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetSyokurekiInfo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetSyokurekiInfo"

            mGetSyokurekiInfo = False

            '�擾
            If fncGetSyokurekiInfo() = False Then
                Exit Function
            End If

            '����I��
            mGetSyokurekiInfo = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��ʊw�������擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetIppanGakurekiInfo() As Boolean

        General.g_ErrorProc = "clsStaffIdo mGetIppanGakurekiInfo"

        mGetIppanGakurekiInfo = False
        Try
            '�擾
            If fncGetIppanGakurekiInfo() = False Then
                Exit Function
            End If

            '����I��
            mGetIppanGakurekiInfo = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���w�������擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetSenmonGakurekiInfo() As Boolean

        General.g_ErrorProc = "clsStaffIdo mGetSenmonGakurekiInfo"

        mGetSenmonGakurekiInfo = False
        Try
            '�擾
            If fncGetSenmonGakurekiInfo() = False Then
                Exit Function
            End If
            '����I��
            mGetSenmonGakurekiInfo = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���x�����擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetChoukyuInfo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetChoukyuInfo"

            mGetChoukyuInfo = False

            '�擾
            If fncGetChoukyuInfo() = False Then
                Exit Function
            End If

            '����I��
            mGetChoukyuInfo = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Y�x�����擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetSankyuInfo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetSankyuInfo"

            mGetSankyuInfo = False

            '�擾
            If fncGetSankyuInfo() = False Then
                Exit Function
            End If

            '����I��
            mGetSankyuInfo = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��������擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetKyoukaiInfo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetKyoukaiInfo"

            mGetKyoukaiInfo = False

            '�擾
            If fncGetKyoukaiInfo() = False Then
                Exit Function
            End If

            '����I��
            mGetKyoukaiInfo = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ƒ������擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetKazokuInfo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetKazokuInfo"

            mGetKazokuInfo = False

            '�擾
            If fncGetKazokuInfo() = False Then
                Exit Function
            End If

            '����I��
            mGetKazokuInfo = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ɛя����擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetGyosekiInfo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetGyosekiInfo"

            mGetGyosekiInfo = False

            '�擾
            If fncGetGyosekiInfo() = False Then
                Exit Function
            End If

            '����I��
            mGetGyosekiInfo = True

        Catch ex As Exception
            Throw
        End Try
    End Function
    '2012/02/13 Sasaki add start-------------------------------------------------------------------------------------------
    '************************���C��u���擾*******************************************
    Public Function mGetStudyAttend() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetStudyAttend"

            mGetStudyAttend = False

            '�擾
            If fncGetStudyAttend() = False Then
                Exit Function
            End If

            '����I��
            mGetStudyAttend = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************DLĻݸ��� f���C��u������*****************************
    Public Function fSA_Count() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_Count"

            fSA_Count = m_lngSACount

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�����è��i�擾�����j*********************
    Public WriteOnly Property mSA_Idx() As Long
        Set(ByVal Value As Long)
            Try
                General.g_ErrorProc = "clsStaffIdo mSA_Idx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_lngSACount Then
                    Exit Property
                End If
                m_lngSAIdx = Value

            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    '************************�i�N�x�j*********************
    Public Function fSA_Nendo() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_Nendo"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_Nendo = 0
            Else
                fSA_Nendo = g_StudyAttend(m_lngSAIdx).lngNendo
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i���C�h�m�c�d�w�j*********************
    Public Function fSA_StudyIndex() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_StudyIndex"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_StudyIndex = 0
            Else
                fSA_StudyIndex = g_StudyAttend(m_lngSAIdx).lngStudyIndex
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�@���O�t���O�j*********************
    Public Function fSA_OutInFlg() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_OutInFlg"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_OutInFlg = ""
            Else
                fSA_OutInFlg = g_StudyAttend(m_lngSAIdx).strOutInFlg
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i���C�R�[�h�j*********************
    Public Function fSA_StudyCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_StudyCD"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_StudyCD = ""
            Else
                fSA_StudyCD = g_StudyAttend(m_lngSAIdx).strStudyCD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i���C���́j*********************
    Public Function fSA_StudyName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_StudyName"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_StudyName = ""
            Else
                fSA_StudyName = g_StudyAttend(m_lngSAIdx).strStudyName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i���C���́j*********************
    Public Function fSA_StudySecName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_StudySecName"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_StudySecName = ""
            Else
                fSA_StudySecName = g_StudyAttend(m_lngSAIdx).strStudySecName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i���C�t���K�i�j*********************
    Public Function fSA_StudyKana() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_StudyKana"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_StudyKana = ""
            Else
                fSA_StudyKana = g_StudyAttend(m_lngSAIdx).strStudyKana
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i��ރR�[�h�j*********************
    Public Function fSA_KindCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_KindCD"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_KindCD = ""
            Else
                fSA_KindCD = g_StudyAttend(m_lngSAIdx).strKindCD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i��ޖ��́j*********************
    Public Function fSA_KindName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_KindName"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_KindName = ""
            Else
                fSA_KindName = g_StudyAttend(m_lngSAIdx).strKindName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i��ÃR�[�h�j*********************
    Public Function fSA_SponsorCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_SponsorCD"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_SponsorCD = ""
            Else
                fSA_SponsorCD = g_StudyAttend(m_lngSAIdx).strSponsorCD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i��Ö��́j*********************
    Public Function fSA_SponsorName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_SponsorName"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_SponsorName = ""
            Else
                fSA_SponsorName = g_StudyAttend(m_lngSAIdx).strSponsorName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�e�[�}�j*********************
    Public Function fSA_Theme() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_Theme"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_Theme = ""
            Else
                fSA_Theme = g_StudyAttend(m_lngSAIdx).strTheme
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�u�t�j*********************
    Public Function fSA_Lecturer() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_Lecturer"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_Lecturer = ""
            Else
                fSA_Lecturer = g_StudyAttend(m_lngSAIdx).strLecturer
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i���E�ꏊ�j*********************
    Public Function fSA_Hall() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_Hall"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_Hall = ""
            Else
                fSA_Hall = g_StudyAttend(m_lngSAIdx).strHall
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�Q�������j*********************
    Public Function fSA_SankaCond() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_SankaCond"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_SankaCond = ""
            Else
                fSA_SankaCond = g_StudyAttend(m_lngSAIdx).strSankaCond
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�Q���l���j*********************
    Public Function fSA_SankaNinzu() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_SankaNinzu"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_SankaNinzu = 0
            Else
                fSA_SankaNinzu = g_StudyAttend(m_lngSAIdx).lngSankaNinzu
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�A�������j*********************
    Public Function fSA_Reports() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_Reports"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_Reports = ""
            Else
                fSA_Reports = g_StudyAttend(m_lngSAIdx).strReports
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i���l�j*********************
    Public Function fSA_Bikou() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_Bikou"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_Bikou = ""
            Else
                fSA_Bikou = g_StudyAttend(m_lngSAIdx).strBikou
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�t�q�k�j*********************
    Public Function fSA_Url() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_Url"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_Url = ""
            Else
                fSA_Url = g_StudyAttend(m_lngSAIdx).strUrl
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�K�{�]�����x���R�[�h�j*********************
    Public Function fSA_NecessaryValuationLevelCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_NecessaryValuationLevelCD"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_NecessaryValuationLevelCD = ""
            Else
                fSA_NecessaryValuationLevelCD = g_StudyAttend(m_lngSAIdx).strNecessaryValuationLevelCD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�K�{�]�����x�����́j*********************
    Public Function fSA_NecessaryValuationLevelName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_NecessaryValuationLevelName"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_NecessaryValuationLevelName = ""
            Else
                fSA_NecessaryValuationLevelName = g_StudyAttend(m_lngSAIdx).strNecessaryValuationLevelName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�K�{�]�����x�����́j*********************
    Public Function fSA_NecessaryValuationLevelSecName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_NecessaryValuationLevelSecName"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_NecessaryValuationLevelSecName = ""
            Else
                fSA_NecessaryValuationLevelSecName = g_StudyAttend(m_lngSAIdx).strNecessaryValuationLevelSecName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�K�{�]�����x���L���j*********************
    Public Function fSA_NecessaryValuationLevelMark() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_NecessaryValuationLevelMark"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_NecessaryValuationLevelMark = ""
            Else
                fSA_NecessaryValuationLevelMark = g_StudyAttend(m_lngSAIdx).strNecessaryValuationLevelMark
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i��t�J�n�N�����j*********************
    Public Function fSA_AcceptFromDate() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_AcceptFromDate"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_AcceptFromDate = 0
            Else
                fSA_AcceptFromDate = g_StudyAttend(m_lngSAIdx).lngAcceptFromDate
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i��t�I���N�����j*********************
    Public Function fSA_AcceptToDate() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_AcceptToDate"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_AcceptToDate = 0
            Else
                fSA_AcceptToDate = g_StudyAttend(m_lngSAIdx).lngAcceptToDate
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i��t�\���݃t���O�j*********************
    Public Function fSA_Acceptapstate() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_Acceptapstate"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_Acceptapstate = ""
            Else
                fSA_Acceptapstate = g_StudyAttend(m_lngSAIdx).strAcceptapstate
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�N�Ԍv��敪�j*********************
    Public Function fSA_NendoPlanKbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_NendoPlanKbn"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_NendoPlanKbn = ""
            Else
                fSA_NendoPlanKbn = g_StudyAttend(m_lngSAIdx).strNendoPlanKbn
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�Ζ������R�[�h�j*********************
    Public Function fSA_KinmuDeptCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_KinmuDeptCD"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_KinmuDeptCD = ""
            Else
                fSA_KinmuDeptCD = g_StudyAttend(m_lngSAIdx).strKinmuDeptCD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�Ζ��������́j*********************
    Public Function fSA_KinmuDeptName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_KinmuDeptName"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_KinmuDeptName = ""
            Else
                fSA_KinmuDeptName = g_StudyAttend(m_lngSAIdx).strKinmuDeptName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�S�����K�{�t���O�j*********************
    Public Function fSA_AllDaysNecessaryFlg() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_AllDaysNecessaryFlg"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_AllDaysNecessaryFlg = ""
            Else
                fSA_AllDaysNecessaryFlg = g_StudyAttend(m_lngSAIdx).strSponsorName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i���匤�C�t���O�j*********************
    Public Function fSA_IndependentFlg() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_IndependentFlg"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_IndependentFlg = ""
            Else
                fSA_IndependentFlg = g_StudyAttend(m_lngSAIdx).strIndependentFlg
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i���t�C���f�b�N�X�j*********************
    Public Function fSA_DateIdx() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_DateIdx"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_DateIdx = 0
            Else
                fSA_DateIdx = g_StudyAttend(m_lngSAIdx).lngDateIdx
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i���t�w��t���O�j*********************
    Public Function fSA_DateAppoFlg() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_DateAppoFlg"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_DateAppoFlg = ""
            Else
                fSA_DateAppoFlg = g_StudyAttend(m_lngSAIdx).strSponsorName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i���t�J�n�N�����j*********************
    Public Function fSA_DateFrom() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_DateFrom"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_DateFrom = 0
            Else
                fSA_DateFrom = g_StudyAttend(m_lngSAIdx).lngDateFrom
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i���t�I���N�����j*********************
    Public Function fSA_DateTo() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_DateTo"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_DateTo = 0
            Else
                fSA_DateTo = g_StudyAttend(m_lngSAIdx).lngDateTo
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�s���{���R�[�h�j*********************
    Public Function fSA_JapanAreaCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_JapanAreaCD"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_JapanAreaCD = ""
            Else
                fSA_JapanAreaCD = g_StudyAttend(m_lngSAIdx).strJapanAreaCD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�s���{�����́j*********************
    Public Function fSA_JapanAreaName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_JapanAreaName"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_JapanAreaName = ""
            Else
                fSA_JapanAreaName = g_StudyAttend(m_lngSAIdx).strJapanAreaName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i����o�^�����j*********************
    Public Function fSA_RegistFirstTimeDate() As Double
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_RegistFirstTimeDate"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_RegistFirstTimeDate = 0
            Else
                fSA_RegistFirstTimeDate = g_StudyAttend(m_lngSAIdx).dblRegistFirstTimeDate
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�ŏI�X�V�����j*********************
    Public Function fSA_LastUpdTimeDate() As Double
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_LastUpdTimeDate"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_LastUpdTimeDate = 0
            Else
                fSA_LastUpdTimeDate = g_StudyAttend(m_lngSAIdx).strSponsorName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�o�^�҂h�c�j*********************
    Public Function fSA_RegistrantID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_RegistrantID"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_RegistrantID = ""
            Else
                fSA_RegistrantID = g_StudyAttend(m_lngSAIdx).strRegistrantID
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i��u�σt���O�j*********************
    Public Function fSA_AttendCompFlg() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_AttendCompFlg"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_AttendCompFlg = ""
            Else
                fSA_AttendCompFlg = g_StudyAttend(m_lngSAIdx).strAttendCompFlg
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i��u�񍐁j*********************
    Public Function fSA_AttendLecrep() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_AttendLecrep"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_AttendLecrep = ""
            Else
                fSA_AttendLecrep = g_StudyAttend(m_lngSAIdx).strAttendLecrep
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i��p�R�[�h�j*********************
    Public Function fSA_CostCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_CostCD"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_CostCD = ""
            Else
                fSA_CostCD = g_StudyAttend(m_lngSAIdx).strCostCD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i��p���́j*********************
    Public Function fSA_CostName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_CostName"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_CostName = ""
            Else
                fSA_CostName = g_StudyAttend(m_lngSAIdx).strCostName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�Q���`�ԃR�[�h�j*********************
    Public Function fSA_SankaFormCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_SankaFormCD"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_SankaFormCD = ""
            Else
                fSA_SankaFormCD = g_StudyAttend(m_lngSAIdx).strSankaFormCD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�Q���`�Ԗ��́j*********************
    Public Function fSA_SankaFormName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_SankaFormName"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_SankaFormName = ""
            Else
                fSA_SankaFormName = g_StudyAttend(m_lngSAIdx).strSankaFormName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i���l�j*********************
    Public Function fSA_SSBikou() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_SSBikou"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_SSBikou = ""
            Else
                fSA_SSBikou = g_StudyAttend(m_lngSAIdx).strSSBikou
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�iUniqueSeqNo�j*********************
    Public Function fSA_UniqueSeqNo() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_UniqueSeqNo"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_UniqueSeqNo = ""
            Else
                fSA_UniqueSeqNo = g_StudyAttend(m_lngSAIdx).strUniqueSeqNo
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i���F�σt���O�j*********************
    Public Function fSA_ApproveFlg() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_ApproveFlg"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_ApproveFlg = ""
            Else
                fSA_ApproveFlg = g_StudyAttend(m_lngSAIdx).strApproveFlg
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '************************�i�Q���t���O�j*********************
    Public Function fSA_SankaFlg() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSA_SankaFlg"

            If m_lngSACount = 0 Or m_lngSAIdx = 0 Then
                fSA_SankaFlg = ""
            Else
                fSA_SankaFlg = g_StudyAttend(m_lngSAIdx).strSankaFlg
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function
    '2012/02/13 Sasaki add end---------------------------------------------------------------------------------------------

    ''' <summary>
    ''' �̗p�ٓ��������Z�b�g����
    ''' </summary>
    ''' <param name="Value">�̗p�ٓ�����</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mSI_SaiyoIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mSI_SaiyoIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numSaiyoKensu Then
                    Exit Property
                End If
                m_numSaiyoIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property


    ''' <summary>
    ''' �Ζ������ٓ��������Z�b�g����
    ''' </summary>
    ''' <param name="Value">�Ζ������ٓ�����</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mKI_KinmuDeptIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mKI_SaiyoIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numKinmuDeptKensu Then
                    Exit Property
                End If
                m_numKinmuDeptIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property
    
    ''' <summary>
    ''' �Ζ������ٓ��������Z�b�g����
    ''' </summary>
    ''' <param name="Value">�Ζ������ٓ�����</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mSAI_KinmuDeptIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mSAI_KinmuDeptIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numSaikeiKensu Then
                    Exit Property
                End If
                m_numSaikeiIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �z�������ٓ��������Z�b�g����
    ''' </summary>
    ''' <param name="Value">�z�������ٓ�����</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mWI_WardDeptIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mWI_WardDeptIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numWardDeptKensu Then
                    Exit Property
                End If
                m_numWardDeptIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' ��E�ٓ��������Z�b�g����
    ''' </summary>
    ''' <param name="Value">��E�ٓ�����</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mPI_PostIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mPI_PostIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numPostKensu Then
                    Exit Property
                End If
                m_numPostIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �E��ٓ��������Z�b�g����
    ''' </summary>
    ''' <param name="Value">�E��ٓ�����</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mJI_JobIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mJI_JobIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numJobKensu Then
                    Exit Property
                End If
                m_numJobIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �����ٓ��������Z�b�g����
    ''' </summary>
    ''' <param name="Value">�����ٓ�����</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mKE_KenmuIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mKE_KenmuIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numKenmuKensu Then
                    Exit Property
                End If
                m_numKenmuIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �Ƌ����������Z�b�g����
    ''' </summary>
    ''' <param name="Value">�Ƌ�������</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mLI_MenkyoIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mLI_MenkyoIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numMenkyoKensu Then
                    Exit Property
                End If
                m_numMenkyoIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' ���i���������Z�b�g����
    ''' </summary>
    ''' <param name="Value">���i������</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mSH_ShikakuIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mSH_ShikakuIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numShikakuKensu Then
                    Exit Property
                End If
                m_numShikakuIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �ψ����������Z�b�g����
    ''' </summary>
    ''' <param name="Value">�ψ�������</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mII_IinIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mII_IinIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numIinKensu Then
                    Exit Property
                End If
                m_numIinIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �E�����������Z�b�g����
    ''' </summary>
    ''' <param name="Value">�E��������</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mJC_SyokurekiIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mJC_SyokurekiIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numSyokurekiKensu Then
                    Exit Property
                End If
                m_numSyokurekiIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property


    ''' <summary>
    ''' ��ʊw�����������Z�b�g����
    ''' </summary>
    ''' <param name="Value">��ʊw��������</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mGS_IppanGakurekiIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mGS_IppanGakurekiIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numIppanGakurekiKensu Then
                    Exit Property
                End If
                m_numIppanGakurekiIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' ���w�����������Z�b�g����
    ''' </summary>
    ''' <param name="Value">���w��������</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mSS_SenmonGakurekiIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mSS_SenmonGakurekiIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numSenmonGakurekiKensu Then
                    Exit Property
                End If
                m_numSenmonGakurekiIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' ���x���������Z�b�g����
    ''' </summary>
    ''' <param name="Value">���x������</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mLL_ChoukyuIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mLL_ChoukyuIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numChoukyuKensu Then
                    Exit Property
                End If
                m_numChoukyuIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �Y�x���������Z�b�g����
    ''' </summary>
    ''' <param name="Value">�Y�x������</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mSK_SankyuIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mSK_SankyuIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numSankyuKensu Then
                    Exit Property
                End If
                m_numSankyuIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �Y�x���������Z�b�g����
    ''' </summary>
    ''' <param name="Value">�Y�x������</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mSO_KyoukaiIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mSO_KyoukaiIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numKyoukaiKensu Then
                    Exit Property
                End If
                m_numKyoukaiIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �Ƒ����������Z�b�g����
    ''' </summary>
    ''' <param name="Value">�Ƒ�������</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mKY_KazokuIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mKY_KazokuIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numKazokuKensu Then
                    Exit Property
                End If
                m_numKazokuIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property


    ''' <summary>
    ''' �a�@CD���Z�b�g����
    ''' </summary>
    ''' <param name="Value">�a�@CD</param>
    ''' <remarks></remarks>
    Public WriteOnly Property pHospitalCD() As String
        Set(ByVal Value As String)
            Try
                General.g_ErrorProc = "clsStaffIdo pHospitalCD"


                m_strHospitalCD = IIf(IsNothing(Value), "", Trim(Value))

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �E���Ǘ��ԍ����Z�b�g����
    ''' </summary>
    ''' <param name="Value">�E���Ǘ��ԍ�</param>
    ''' <remarks></remarks>
    Public WriteOnly Property pStaffMngID() As String
        Set(ByVal Value As String)
            Try
                General.g_ErrorProc = "clsStaffIdo pStaffMngID"


                m_strStaffMngID = IIf(IsNothing(Value), "", Trim(Value))

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' ���t�敪���Z�b�g����
    ''' </summary>
    ''' <param name="Value">���t�敪</param>
    ''' <remarks></remarks>
    Public WriteOnly Property pDateFlg() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo pDateFlg"

                '0:�P��� 1:���� 2:�S��
                m_numDateFlg = IIf(IsNothing(Value), 0, Value)

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �J�n�N�������Z�b�g����
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <remarks></remarks>
    Public WriteOnly Property pDateFrom() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo pDateFrom"

                m_numDateFrom = IIf(IsNothing(Value), 0, Value)

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �I���N�������Z�b�g����
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <remarks></remarks>
    Public WriteOnly Property pDateTo() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo pDateTo"

                m_numDateTo = IIf(IsNothing(Value), 0, Value)

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �\�[�g�����Z�b�g����
    ''' </summary>
    ''' <param name="Value">�\�[�g��</param>
    ''' <remarks></remarks>
    Public WriteOnly Property pSortFlg() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo pSortFlg"

                m_numSortFlg = IIf(IsNothing(Value), 0, Value)

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �N�x���Z�b�g����
    ''' </summary>
    ''' <param name="Value">�N�x</param>
    ''' <remarks></remarks>
    Public WriteOnly Property pNendo() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo pNendo"

                m_numNendo = IIf(IsNothing(Value), 0, Value)

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �폜�󋵂��Z�b�g����
    ''' </summary>
    ''' <param name="Value">�폜��</param>
    ''' <remarks></remarks>
    Public WriteOnly Property pDelKbn() As Short
        Set(ByVal Value As Short)
            Try
                General.g_ErrorProc = "clsStaffIdo pDelKbn"

                m_intDelKbn = IIf(IsNothing(Value), 2, Value)

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property
    '2012/02/13 Sasaki add start--------------------------------------------------------------------------------
    '************************DLL�K�{�����è��i�Q���t���O�j*********************
    Public WriteOnly Property pSankaFlg() As String
        Set(ByVal Value As String)
            Try
                General.g_ErrorProc = "clsStaffIdo pSankaFlg"

                m_strSankaFlg = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    '************************DLL�K�{�����è��i�擾�N����FROM�j*********************
    Public WriteOnly Property pGetFromDate() As Long
        Set(ByVal Value As Long)
            Try
                General.g_ErrorProc = "clsStaffIdo pGetFromDate"

                m_numDateFrom = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    '************************DLL�K�{�����è��i�擾�N����TO�j*********************
    Public WriteOnly Property pGetToDate() As Long
        Set(ByVal Value As Long)
            Try
                General.g_ErrorProc = "clsStaffIdo pGetToDate"

                m_numDateTo = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    '************************DLL�K�{�����è��i���C�h�m�c�d�w�j*********************
    Public WriteOnly Property pStudyIndex() As Long
        Set(ByVal Value As Long)
            Try
                General.g_ErrorProc = "clsStaffIdo pStudyIndex"

                m_numStudyIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    '************************DLL�K�{�����è��i��u�σt���O�j*********************
    Public WriteOnly Property pAttendCompFlg() As String
        Set(ByVal Value As String)
            Try
                General.g_ErrorProc = "clsStaffIdo pAttendCompFlg"

                m_strAttendCompFlg = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    '************************DLL�K�{�����è��i���F�σt���O�j*********************
    Public WriteOnly Property pApproveFlg() As String
        Set(ByVal Value As String)
            Try
                General.g_ErrorProc = "clsStaffIdo pApproveFlg"

                m_strApproveFlg = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    '************************DLL�K�{�����è��i�@���O�t���O�j*********************
    Public WriteOnly Property pOutInFlg() As String
        Set(ByVal Value As String)
            Try
                General.g_ErrorProc = "clsStaffIdo pOutInFlg"

                m_strOutInFlg = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property
    '2012/02/13 Sasaki add end----------------------------------------------------------------------------------

    ''' <summary>
    ''' ���C���������Z�b�g����
    ''' </summary>
    ''' <param name="Value">���C������</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mSD_StudyIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mSD_StudyIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numStudyKensu Then
                    Exit Property
                End If
                m_numStudyIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' ���C���t���������Z�b�g����
    ''' </summary>
    ''' <param name="Value">���C���t������</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mSD_StudyDateIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mSD_StudyDateIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numStudyDateKensu Then
                    Exit Property
                End If
                m_numStudyDateIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property
    ''' <summary>
    ''' �Ɛя��������Z�b�g����
    ''' </summary>
    ''' <param name="Value">�Ɛя�����</param>
    ''' <remarks></remarks>
    Public WriteOnly Property mGY_GyosekiIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mGY_GyosekiIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numGyosekiKensu Then
                    Exit Property
                End If
                m_numGyosekiIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �̗p�ٓ��������擾����
    ''' </summary>
    ''' <returns>�̗p�ٓ�����</returns>
    ''' <remarks></remarks>
    Public Function fSI_SaiyoCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSI_SaiyoCount"

            fSI_SaiyoCount = m_numSaiyoKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>�̗p�ٓ������C���f�b�N�X���O�A�܂��́A�̗p�ٓ������������O�̏ꍇ�A""���擾����</remarks>
    Public Function fSI_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSI_HospitalCD"

            If m_numSaiyoIdx = 0 OrElse m_numSaiyoKensu = 0 Then
                fSI_HospitalCD = ""
            Else
                fSI_HospitalCD = g_SaiyoIdo.strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>�̗p�ٓ������C���f�b�N�X���O�A�܂��́A�̗p�ٓ������������O�̏ꍇ�A""���擾����</remarks>
    Public Function fSI_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSI_StaffMngID"

            If m_numSaiyoIdx = 0 OrElse m_numSaiyoKensu = 0 Then
                fSI_StaffMngID = ""
            Else
                fSI_StaffMngID = g_SaiyoIdo.strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �̗p�R�[�h���擾����
    ''' </summary>
    ''' <returns>�̗p�R�[�h</returns>
    ''' <remarks>�̗p�ٓ������C���f�b�N�X���O�A�܂��́A�̗p�ٓ������������O�̏ꍇ�A""���擾����</remarks>
    Public Function fSI_EmpCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSI_EmpCD"

            If m_numSaiyoIdx = 0 OrElse m_numSaiyoKensu = 0 Then
                fSI_EmpCD = ""
            Else
                fSI_EmpCD = g_SaiyoIdo.Ido(m_numSaiyoIdx).strEmpCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �̗p���̂��擾����
    ''' </summary>
    ''' <returns>�̗p����</returns>
    ''' <remarks>�̗p�ٓ������C���f�b�N�X���O�A�܂��́A�̗p�ٓ������������O�̏ꍇ�A""���擾����</remarks>
    Public Function fSI_EmpName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSI_EmpName"

            If m_numSaiyoIdx = 0 OrElse m_numSaiyoKensu = 0 Then
                fSI_EmpName = ""
            Else
                fSI_EmpName = g_SaiyoIdo.Ido(m_numSaiyoIdx).strEmpName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �̗p���̂��擾����
    ''' </summary>
    ''' <returns>�̗p����</returns>
    ''' <remarks>�̗p�ٓ������C���f�b�N�X���O�A�܂��́A�̗p�ٓ������������O�̏ꍇ�A""���擾����</remarks>
    Public Function fSI_EmpSecName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSI_EmpName"

            If m_numSaiyoIdx = 0 OrElse m_numSaiyoKensu = 0 Then
                fSI_EmpSecName = ""
            Else
                fSI_EmpSecName = g_SaiyoIdo.Ido(m_numSaiyoIdx).strEmpSecName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �̗p���t���擾����
    ''' </summary>
    ''' <returns>�̗p���t</returns>
    ''' <remarks>�̗p�ٓ������C���f�b�N�X���O�A�܂��́A�̗p�ٓ������������O�̏ꍇ�A0���擾����</remarks>
    Public Function fSI_EmpDate() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSI_EmpDate"

            If m_numSaiyoIdx = 0 OrElse m_numSaiyoKensu = 0 Then
                fSI_EmpDate = 0
            Else
                fSI_EmpDate = g_SaiyoIdo.Ido(m_numSaiyoIdx).numEmpDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ސE�R�[�h���擾����
    ''' </summary>
    ''' <returns>�ސE�R�[�h</returns>
    ''' <remarks>�̗p�ٓ������C���f�b�N�X���O�A�܂��́A�̗p�ٓ������������O�̏ꍇ�A""���擾����</remarks>
    Public Function fSI_RetireCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSI_RetireCD"

            If m_numSaiyoIdx = 0 OrElse m_numSaiyoKensu = 0 Then
                fSI_RetireCD = ""
            Else
                fSI_RetireCD = g_SaiyoIdo.Ido(m_numSaiyoIdx).strRetireCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ސE���̂��擾����
    ''' </summary>
    ''' <returns>�ސE����</returns>
    ''' <remarks>�̗p�ٓ������C���f�b�N�X���O�A�܂��́A�̗p�ٓ������������O�̏ꍇ�A""���擾����</remarks>
    Public Function fSI_RetireName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSI_RetireName"

            If m_numSaiyoIdx = 0 OrElse m_numSaiyoKensu = 0 Then
                fSI_RetireName = ""
            Else
                fSI_RetireName = g_SaiyoIdo.Ido(m_numSaiyoIdx).strRetireName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ސE���t���擾����
    ''' </summary>
    ''' <returns>�ސE���t</returns>
    ''' <remarks>�̗p�ٓ������C���f�b�N�X���O�A�܂��́A�̗p�ٓ������������O�̏ꍇ�A0���擾����</remarks>
    Public Function fSI_RetireDate() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSI_RetireDate"

            If m_numSaiyoIdx = 0 OrElse m_numSaiyoKensu = 0 Then
                fSI_RetireDate = 0
            Else
                fSI_RetireDate = g_SaiyoIdo.Ido(m_numSaiyoIdx).numRetireDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>�̗p�ٓ������C���f�b�N�X���O�A�܂��́A�̗p�ٓ������������O�̏ꍇ�A0���擾����</remarks>
    Public Function fSI_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSI_FirstTime"

            If m_numSaiyoIdx = 0 OrElse m_numSaiyoKensu = 0 Then
                fSI_FirstTime = 0
            Else
                fSI_FirstTime = g_SaiyoIdo.Ido(m_numSaiyoIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>�̗p�ٓ������C���f�b�N�X���O�A�܂��́A�̗p�ٓ������������O�̏ꍇ�A0���擾����</remarks>
    Public Function fSI_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSI_UpdTime"

            If m_numSaiyoIdx = 0 OrElse m_numSaiyoKensu = 0 Then
                fSI_UpdTime = 0
            Else
                fSI_UpdTime = g_SaiyoIdo.Ido(m_numSaiyoIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' �E���ԍ����擾����
    ''' </summary>
    ''' <returns>�E���ԍ�</returns>
    ''' <remarks>�̗p�ٓ������C���f�b�N�X���O�A�܂��́A�̗p�ٓ������������O�̏ꍇ�A""���擾����</remarks>
    Public Function fSI_StaffID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSI_StaffID"

            If m_numSaiyoIdx = 0 OrElse m_numSaiyoKensu = 0 Then
                fSI_StaffID = ""
            Else
                fSI_StaffID = g_SaiyoIdo.Ido(m_numSaiyoIdx).strStaffID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ζ������ٓ��������擾����
    ''' </summary>
    ''' <returns>�Ζ������ٓ�����</returns>
    ''' <remarks></remarks>
    Public Function fKI_KinmuDeptCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fKI_SaiyoCount"

            fKI_KinmuDeptCount = m_numKinmuDeptKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>�Ζ������ٓ��C���f�b�N�X���O�A�܂��́A�Ζ������ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fKI_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKI_HospitalCD"

            If m_numKinmuDeptIdx = 0 OrElse m_numKinmuDeptKensu = 0 Then
                fKI_HospitalCD = ""
            Else
                fKI_HospitalCD = g_KinmuDeptIdo.strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>�Ζ������ٓ��C���f�b�N�X���O�A�܂��́A�Ζ������ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fKI_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKI_StaffMngID"

            If m_numKinmuDeptIdx = 0 OrElse m_numKinmuDeptKensu = 0 Then
                fKI_StaffMngID = ""
            Else
                fKI_StaffMngID = g_KinmuDeptIdo.strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �R�[�h���擾����
    ''' </summary>
    ''' <returns>�R�[�h</returns>
    ''' <remarks>�Ζ������ٓ��C���f�b�N�X���O�A�܂��́A�Ζ������ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fKI_CD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKI_CD"

            If m_numKinmuDeptIdx = 0 OrElse m_numKinmuDeptKensu = 0 Then
                fKI_CD = ""
            Else
                fKI_CD = g_KinmuDeptIdo.Ido(m_numKinmuDeptIdx).strCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���̂��擾����
    ''' </summary>
    ''' <returns>����</returns>
    ''' <remarks>�Ζ������ٓ��C���f�b�N�X���O�A�܂��́A�Ζ������ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fKI_Name() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKI_Name"

            If m_numKinmuDeptIdx = 0 OrElse m_numKinmuDeptKensu = 0 Then
                fKI_Name = ""
            Else
                fKI_Name = g_KinmuDeptIdo.Ido(m_numKinmuDeptIdx).strName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �J�n�����擾����
    ''' </summary>
    ''' <returns>�J�n��</returns>
    ''' <remarks>�Ζ������ٓ��C���f�b�N�X���O�A�܂��́A�Ζ������ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fKI_DateFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fKI_DateFrom"

            If m_numKinmuDeptIdx = 0 OrElse m_numKinmuDeptKensu = 0 Then
                fKI_DateFrom = 0
            Else
                fKI_DateFrom = g_KinmuDeptIdo.Ido(m_numKinmuDeptIdx).numDateFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �I�������擾����
    ''' </summary>
    ''' <returns>�I����</returns>
    ''' <remarks>�Ζ������ٓ��C���f�b�N�X���O�A�܂��́A�Ζ������ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fKI_DateTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fKI_DateTo"

            If m_numKinmuDeptIdx = 0 OrElse m_numKinmuDeptKensu = 0 Then
                fKI_DateTo = 0
            Else
                fKI_DateTo = g_KinmuDeptIdo.Ido(m_numKinmuDeptIdx).numDateTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>�Ζ������ٓ��C���f�b�N�X���O�A�܂��́A�Ζ������ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fKI_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fKI_FirstTime"

            If m_numKinmuDeptIdx = 0 OrElse m_numKinmuDeptKensu = 0 Then
                fKI_FirstTime = 0
            Else
                fKI_FirstTime = g_KinmuDeptIdo.Ido(m_numKinmuDeptIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>�Ζ������ٓ��C���f�b�N�X���O�A�܂��́A�Ζ������ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fKI_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fKI_UpdTime"

            If m_numKinmuDeptIdx = 0 OrElse m_numKinmuDeptKensu = 0 Then
                fKI_UpdTime = 0
            Else
                fKI_UpdTime = g_KinmuDeptIdo.Ido(m_numKinmuDeptIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' �ٓ���]���擾����
    ''' </summary>
    ''' <returns>�ٓ���]</returns>
    ''' <remarks></remarks>
    Public Function fKI_IdoHope() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKI_IdoHope"

            fKI_IdoHope = g_KinmuDeptIdo.Ido(m_numKinmuDeptIdx).strIdoHope


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���̂��擾����
    ''' </summary>
    ''' <returns>����</returns>
    ''' <remarks></remarks>
    Public Function fKI_SecName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKI_SecName"

            fKI_SecName = g_KinmuDeptIdo.Ido(m_numKinmuDeptIdx).SecName

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �\�������擾����
    ''' </summary>
    ''' <returns>�\����</returns>
    ''' <remarks></remarks>
    Public Function fKI_DispNo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fKI_DispNo"

            fKI_DispNo = g_KinmuDeptIdo.Ido(m_numKinmuDeptIdx).DispNo

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ζ������ٓ��������擾����
    ''' </summary>
    ''' <returns>�Ζ������ٓ�����</returns>
    ''' <remarks></remarks>
    Public Function fSAI_KinmuDeptCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSAI_SaiyoCount"

            fSAI_KinmuDeptCount = m_numSaikeiKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>�Čf�ٓ��C���f�b�N�X���O�A�܂��́A�Čf�ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fSAI_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSAI_HospitalCD"

            If m_numSaikeiIdx = 0 OrElse m_numSaikeiKensu = 0 Then
                fSAI_HospitalCD = ""
            Else
                fSAI_HospitalCD = g_SaikeiIdo.strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>�Čf�ٓ��C���f�b�N�X���O�A�܂��́A�Čf�ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fSAI_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSAI_StaffMngID"

            If m_numSaikeiIdx = 0 OrElse m_numSaikeiKensu = 0 Then
                fSAI_StaffMngID = ""
            Else
                fSAI_StaffMngID = g_SaikeiIdo.strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �R�[�h���擾����
    ''' </summary>
    ''' <returns>�R�[�h</returns>
    ''' <remarks>�Čf�ٓ��C���f�b�N�X���O�A�܂��́A�Čf�ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fSAI_CD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSAI_CD"

            If m_numSaikeiIdx = 0 OrElse m_numSaikeiKensu = 0 Then
                fSAI_CD = ""
            Else
                fSAI_CD = g_SaikeiIdo.Ido(m_numSaikeiIdx).strCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���̂��擾����
    ''' </summary>
    ''' <returns>����</returns>
    ''' <remarks>�Čf�ٓ��C���f�b�N�X���O�A�܂��́A�Čf�ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fSAI_Name() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSAI_Name"

            If m_numSaikeiIdx = 0 OrElse m_numSaikeiKensu = 0 Then
                fSAI_Name = ""
            Else
                fSAI_Name = g_SaikeiIdo.Ido(m_numSaikeiIdx).strName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �J�n�����擾����
    ''' </summary>
    ''' <returns>�J�n��</returns>
    ''' <remarks>�Čf�ٓ��C���f�b�N�X���O�A�܂��́A�Čf�ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fSAI_DateFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSAI_DateFrom"

            If m_numSaikeiIdx = 0 OrElse m_numSaikeiKensu = 0 Then
                fSAI_DateFrom = 0
            Else
                fSAI_DateFrom = g_SaikeiIdo.Ido(m_numSaikeiIdx).numDateFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �I�������擾����
    ''' </summary>
    ''' <returns>�I����</returns>
    ''' <remarks>�Čf�ٓ��C���f�b�N�X���O�A�܂��́A�Čf�ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fSAI_DateTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSAI_DateTo"

            If m_numSaikeiIdx = 0 OrElse m_numSaikeiKensu = 0 Then
                fSAI_DateTo = 0
            Else
                fSAI_DateTo = g_SaikeiIdo.Ido(m_numSaikeiIdx).numDateTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>�Čf�ٓ��C���f�b�N�X���O�A�܂��́A�Čf�ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fSAI_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSAI_FirstTime"

            If m_numSaikeiIdx = 0 OrElse m_numSaikeiKensu = 0 Then
                fSAI_FirstTime = 0
            Else
                fSAI_FirstTime = g_SaikeiIdo.Ido(m_numSaikeiIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>�Čf�ٓ��C���f�b�N�X���O�A�܂��́A�Čf�ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fSAI_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSAI_UpdTime"

            If m_numSaikeiIdx = 0 OrElse m_numSaikeiKensu = 0 Then
                fSAI_UpdTime = 0
            Else
                fSAI_UpdTime = g_SaikeiIdo.Ido(m_numSaikeiIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �z�������ٓ��������擾����
    ''' </summary>
    ''' <returns>�z�������ٓ�����</returns>
    ''' <remarks></remarks>
    Public Function fWI_WardDeptCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fWI_WardDeptCount"

            fWI_WardDeptCount = m_numWardDeptKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>�z�������ٓ��C���f�b�N�X���O�A�܂��́A�z�������ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fWI_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fWI_HospitalCD"

            If m_numWardDeptIdx = 0 OrElse m_numWardDeptKensu = 0 Then
                fWI_HospitalCD = ""
            Else
                fWI_HospitalCD = g_WardDeptIdo.strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>�z�������ٓ��C���f�b�N�X���O�A�܂��́A�z�������ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fWI_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fWI_StaffMngID"

            If m_numWardDeptIdx = 0 OrElse m_numWardDeptKensu = 0 Then
                fWI_StaffMngID = ""
            Else
                fWI_StaffMngID = g_WardDeptIdo.strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �R�[�h���擾����
    ''' </summary>
    ''' <returns>�R�[�h</returns>
    ''' <remarks>�z�������ٓ��C���f�b�N�X���O�A�܂��́A�z�������ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fWI_CD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fWI_CD"

            If m_numWardDeptIdx = 0 OrElse m_numWardDeptKensu = 0 Then
                fWI_CD = ""
            Else
                fWI_CD = g_WardDeptIdo.Ido(m_numWardDeptIdx).strCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���̂��擾����
    ''' </summary>
    ''' <returns>����</returns>
    ''' <remarks>�z�������ٓ��C���f�b�N�X���O�A�܂��́A�z�������ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fWI_Name() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fWI_Name"

            If m_numWardDeptIdx = 0 OrElse m_numWardDeptKensu = 0 Then
                fWI_Name = ""
            Else
                fWI_Name = g_WardDeptIdo.Ido(m_numWardDeptIdx).strName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �J�n�����擾����
    ''' </summary>
    ''' <returns>�J�n��</returns>
    ''' <remarks>�z�������ٓ��C���f�b�N�X���O�A�܂��́A�z�������ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fWI_DateFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fWI_DateFrom"

            If m_numWardDeptIdx = 0 OrElse m_numWardDeptKensu = 0 Then
                fWI_DateFrom = 0
            Else
                fWI_DateFrom = g_WardDeptIdo.Ido(m_numWardDeptIdx).numDateFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �I�������擾����
    ''' </summary>
    ''' <returns>�I����</returns>
    ''' <remarks>�z�������ٓ��C���f�b�N�X���O�A�܂��́A�z�������ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fWI_DateTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fWI_DateTo"

            If m_numWardDeptIdx = 0 OrElse m_numWardDeptKensu = 0 Then
                fWI_DateTo = 0
            Else
                fWI_DateTo = g_WardDeptIdo.Ido(m_numWardDeptIdx).numDateTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>�z�������ٓ��C���f�b�N�X���O�A�܂��́A�z�������ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fWI_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fWI_FirstTime"

            If m_numWardDeptIdx = 0 OrElse m_numWardDeptKensu = 0 Then
                fWI_FirstTime = 0
            Else
                fWI_FirstTime = g_WardDeptIdo.Ido(m_numWardDeptIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>�z�������ٓ��C���f�b�N�X���O�A�܂��́A�z�������ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fWI_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fWI_UpdTime"

            If m_numWardDeptIdx = 0 OrElse m_numWardDeptKensu = 0 Then
                fWI_UpdTime = 0
            Else
                fWI_UpdTime = g_WardDeptIdo.Ido(m_numWardDeptIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��E�ٓ��������擾����
    ''' </summary>
    ''' <returns>��E�ٓ�����</returns>
    ''' <remarks></remarks>
    Public Function fPI_PostCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fPI_PostCount"

            fPI_PostCount = m_numPostKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>��E�ٓ��C���f�b�N�X���O�A�܂��́A��E�ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fPI_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fPI_HospitalCD"

            If m_numPostIdx = 0 OrElse m_numPostKensu = 0 Then
                fPI_HospitalCD = ""
            Else
                fPI_HospitalCD = g_PostIdo.strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>��E�ٓ��C���f�b�N�X���O�A�܂��́A��E�ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fPI_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fPI_StaffMngID"

            If m_numPostIdx = 0 OrElse m_numPostKensu = 0 Then
                fPI_StaffMngID = ""
            Else
                fPI_StaffMngID = g_PostIdo.strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �R�[�h���擾����
    ''' </summary>
    ''' <returns>�R�[�h</returns>
    ''' <remarks>��E�ٓ��C���f�b�N�X���O�A�܂��́A��E�ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fPI_CD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fPI_CD"

            If m_numPostIdx = 0 OrElse m_numPostKensu = 0 Then
                fPI_CD = ""
            Else
                fPI_CD = g_PostIdo.Ido(m_numPostIdx).strCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���̂��擾����
    ''' </summary>
    ''' <returns>����</returns>
    ''' <remarks>��E�ٓ��C���f�b�N�X���O�A�܂��́A��E�ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fPI_Name() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fPI_Name"

            If m_numPostIdx = 0 OrElse m_numPostKensu = 0 Then
                fPI_Name = ""
            Else
                fPI_Name = g_PostIdo.Ido(m_numPostIdx).strName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���̂��擾����
    ''' </summary>
    ''' <returns>����</returns>
    ''' <remarks>��E�ٓ��C���f�b�N�X���O�A�܂��́A��E�ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fPI_SecName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fPI_SecName"

            If m_numPostIdx = 0 OrElse m_numPostKensu = 0 Then
                fPI_SecName = ""
            Else
                fPI_SecName = g_PostIdo.Ido(m_numPostIdx).SecName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �J�n�����擾����
    ''' </summary>
    ''' <returns>�J�n��</returns>
    ''' <remarks>��E�ٓ��C���f�b�N�X���O�A�܂��́A��E�ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fPI_DateFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fPI_DateFrom"

            If m_numPostIdx = 0 OrElse m_numPostKensu = 0 Then
                fPI_DateFrom = 0
            Else
                fPI_DateFrom = g_PostIdo.Ido(m_numPostIdx).numDateFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �I�������擾����
    ''' </summary>
    ''' <returns>�I����</returns>
    ''' <remarks>��E�ٓ��C���f�b�N�X���O�A�܂��́A��E�ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fPI_DateTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fPI_DateTo"

            If m_numPostIdx = 0 OrElse m_numPostKensu = 0 Then
                fPI_DateTo = 0
            Else
                fPI_DateTo = g_PostIdo.Ido(m_numPostIdx).numDateTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>��E�ٓ��C���f�b�N�X���O�A�܂��́A��E�ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fPI_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fPI_FirstTime"

            If m_numPostIdx = 0 OrElse m_numPostKensu = 0 Then
                fPI_FirstTime = 0
            Else
                fPI_FirstTime = g_PostIdo.Ido(m_numPostIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>��E�ٓ��C���f�b�N�X���O�A�܂��́A��E�ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fPI_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fPI_UpdTime"

            If m_numPostIdx = 0 OrElse m_numPostKensu = 0 Then
                fPI_UpdTime = 0
            Else
                fPI_UpdTime = g_PostIdo.Ido(m_numPostIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E��ٓ��������擾����
    ''' </summary>
    ''' <returns>�E��ٓ�����</returns>
    ''' <remarks></remarks>
    Public Function fJI_JobCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fJI_JobCount"

            fJI_JobCount = m_numJobKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>�E��ٓ��C���f�b�N�X���O�A�܂��́A�E��ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fJI_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fJI_HospitalCD"

            If m_numJobIdx = 0 OrElse m_numJobKensu = 0 Then
                fJI_HospitalCD = ""
            Else
                fJI_HospitalCD = g_JobIdo.strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>�E��ٓ��C���f�b�N�X���O�A�܂��́A�E��ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fJI_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fJI_StaffMngID"

            If m_numJobIdx = 0 OrElse m_numJobKensu = 0 Then
                fJI_StaffMngID = ""
            Else
                fJI_StaffMngID = g_JobIdo.strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �R�[�h���擾����
    ''' </summary>
    ''' <returns>�R�[�h</returns>
    ''' <remarks>�E��ٓ��C���f�b�N�X���O�A�܂��́A�E��ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fJI_CD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fJI_CD"

            If m_numJobIdx = 0 OrElse m_numJobKensu = 0 Then
                fJI_CD = ""
            Else
                fJI_CD = g_JobIdo.Ido(m_numJobIdx).strCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���̂��擾����
    ''' </summary>
    ''' <returns>����</returns>
    ''' <remarks>�E��ٓ��C���f�b�N�X���O�A�܂��́A�E��ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fJI_Name() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fJI_Name"

            If m_numJobIdx = 0 OrElse m_numJobKensu = 0 Then
                fJI_Name = ""
            Else
                fJI_Name = g_JobIdo.Ido(m_numJobIdx).strName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���̂��擾����
    ''' </summary>
    ''' <returns>����</returns>
    ''' <remarks>�E��ٓ��C���f�b�N�X���O�A�܂��́A�E��ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fJI_SecName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fJI_SecName"

            If m_numJobIdx = 0 OrElse m_numJobKensu = 0 Then
                fJI_SecName = ""
            Else
                fJI_SecName = g_JobIdo.Ido(m_numJobIdx).SecName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �J�n�����擾����
    ''' </summary>
    ''' <returns>�J�n��</returns>
    ''' <remarks>�E��ٓ��C���f�b�N�X���O�A�܂��́A�E��ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fJI_DateFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fJI_DateFrom"

            If m_numJobIdx = 0 OrElse m_numJobKensu = 0 Then
                fJI_DateFrom = 0
            Else
                fJI_DateFrom = g_JobIdo.Ido(m_numJobIdx).numDateFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �I�������擾����
    ''' </summary>
    ''' <returns>�I����</returns>
    ''' <remarks>�E��ٓ��C���f�b�N�X���O�A�܂��́A�E��ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fJI_DateTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fJI_DateTo"

            If m_numJobIdx = 0 OrElse m_numJobKensu = 0 Then
                fJI_DateTo = 0
            Else
                fJI_DateTo = g_JobIdo.Ido(m_numJobIdx).numDateTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>�E��ٓ��C���f�b�N�X���O�A�܂��́A�E��ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fJI_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fJI_FirstTime"

            If m_numJobIdx = 0 OrElse m_numJobKensu = 0 Then
                fJI_FirstTime = 0
            Else
                fJI_FirstTime = g_JobIdo.Ido(m_numJobIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>�E��ٓ��C���f�b�N�X���O�A�܂��́A�E��ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fJI_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fJI_UpdTime"

            If m_numJobIdx = 0 OrElse m_numJobKensu = 0 Then
                fJI_UpdTime = 0
            Else
                fJI_UpdTime = g_JobIdo.Ido(m_numJobIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �����ٓ��������擾����
    ''' </summary>
    ''' <returns>�����ٓ�����</returns>
    ''' <remarks></remarks>
    Public Function fKE_KenmuCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fKE_KenmuCount"

            fKE_KenmuCount = m_numKenmuKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>�����ٓ��C���f�b�N�X���O�A�܂��́A�����ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fKE_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKE_HospitalCD"

            If m_numKenmuIdx = 0 OrElse m_numKenmuKensu = 0 Then
                fKE_HospitalCD = ""
            Else
                fKE_HospitalCD = g_KenmuIdo.strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>�����ٓ��C���f�b�N�X���O�A�܂��́A�����ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fKE_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKE_StaffMngID"

            If m_numKenmuIdx = 0 OrElse m_numKenmuKensu = 0 Then
                fKE_StaffMngID = ""
            Else
                fKE_StaffMngID = g_KenmuIdo.strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �z�������R�[�h���擾����
    ''' </summary>
    ''' <returns>�z�������R�[�h</returns>
    ''' <remarks>�����ٓ��C���f�b�N�X���O�A�܂��́A�����ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fKE_WardDeptCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKE_WardDeptCD"

            If m_numKenmuIdx = 0 OrElse m_numKenmuKensu = 0 Then
                fKE_WardDeptCD = ""
            Else
                fKE_WardDeptCD = g_KenmuIdo.Ido(m_numKenmuIdx).strWardDeptCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �z���������̂��擾����
    ''' </summary>
    ''' <returns>�z����������</returns>
    ''' <remarks>�����ٓ��C���f�b�N�X���O�A�܂��́A�����ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fKE_WardDeptName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKE_WardDeptName"

            If m_numKenmuIdx = 0 OrElse m_numKenmuKensu = 0 Then
                fKE_WardDeptName = ""
            Else
                fKE_WardDeptName = g_KenmuIdo.Ido(m_numKenmuIdx).strWardDeptName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ζ������R�[�h���擾����
    ''' </summary>
    ''' <returns>�Ζ������R�[�h</returns>
    ''' <remarks>�����ٓ��C���f�b�N�X���O�A�܂��́A�����ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fKE_KinmuDeptCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKE_KinmuDeptCD"

            If m_numKenmuIdx = 0 OrElse m_numKenmuKensu = 0 Then
                fKE_KinmuDeptCD = ""
            Else
                fKE_KinmuDeptCD = g_KenmuIdo.Ido(m_numKenmuIdx).strKinmuDeptCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ζ��������̂��擾����
    ''' </summary>
    ''' <returns>�Ζ���������</returns>
    ''' <remarks>�����ٓ��C���f�b�N�X���O�A�܂��́A�����ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fKE_KinmuDeptName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKE_KinmuDeptName"

            If m_numKenmuIdx = 0 OrElse m_numKenmuKensu = 0 Then
                fKE_KinmuDeptName = ""
            Else
                fKE_KinmuDeptName = g_KenmuIdo.Ido(m_numKenmuIdx).strKinmuDeptName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��E�R�[�h���擾����
    ''' </summary>
    ''' <returns>��E�R�[�h</returns>
    ''' <remarks>�����ٓ��C���f�b�N�X���O�A�܂��́A�����ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fKE_PostCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKE_PostCD"

            If m_numKenmuIdx = 0 OrElse m_numKenmuKensu = 0 Then
                fKE_PostCD = ""
            Else
                fKE_PostCD = g_KenmuIdo.Ido(m_numKenmuIdx).strPostCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��E���̂��擾����
    ''' </summary>
    ''' <returns>��E����</returns>
    ''' <remarks>�����ٓ��C���f�b�N�X���O�A�܂��́A�����ٓ��������O�̏ꍇ�A""���擾����</remarks>
    Public Function fKE_PostName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKE_PostName"

            If m_numKenmuIdx = 0 OrElse m_numKenmuKensu = 0 Then
                fKE_PostName = ""
            Else
                fKE_PostName = g_KenmuIdo.Ido(m_numKenmuIdx).strPostName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �J�n�����擾����
    ''' </summary>
    ''' <returns>�J�n��</returns>
    ''' <remarks>�����ٓ��C���f�b�N�X���O�A�܂��́A�����ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fKE_DateFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fKE_DateFrom"

            If m_numKenmuIdx = 0 OrElse m_numKenmuKensu = 0 Then
                fKE_DateFrom = 0
            Else
                fKE_DateFrom = g_KenmuIdo.Ido(m_numKenmuIdx).numDateFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' SEQ���擾����
    ''' </summary>
    ''' <returns>SEQ</returns>
    ''' <remarks>�����ٓ��C���f�b�N�X���O�A�܂��́A�����ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fKE_SEQ() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fKE_SEQ"

            If m_numKenmuIdx = 0 OrElse m_numKenmuKensu = 0 Then
                fKE_SEQ = 0
            Else
                fKE_SEQ = g_KenmuIdo.Ido(m_numKenmuIdx).numSEQ
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �I�������擾����
    ''' </summary>
    ''' <returns>�I����</returns>
    ''' <remarks>�����ٓ��C���f�b�N�X���O�A�܂��́A�����ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fKE_DateTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fKE_DateTo"

            If m_numKenmuIdx = 0 OrElse m_numKenmuKensu = 0 Then
                fKE_DateTo = 0
            Else
                fKE_DateTo = g_KenmuIdo.Ido(m_numKenmuIdx).numDateTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>�����ٓ��C���f�b�N�X���O�A�܂��́A�����ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fKE_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fKE_FirstTime"

            If m_numKenmuIdx = 0 OrElse m_numKenmuKensu = 0 Then
                fKE_FirstTime = 0
            Else
                fKE_FirstTime = g_KenmuIdo.Ido(m_numKenmuIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>�����ٓ��C���f�b�N�X���O�A�܂��́A�����ٓ��������O�̏ꍇ�A0���擾����</remarks>
    Public Function fKE_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fKE_UpdTime"

            If m_numKenmuIdx = 0 OrElse m_numKenmuKensu = 0 Then
                fKE_UpdTime = 0
            Else
                fKE_UpdTime = g_KenmuIdo.Ido(m_numKenmuIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ƌ���񌏐����擾����
    ''' </summary>
    ''' <returns>�Ƌ���񌏐�</returns>
    ''' <remarks></remarks>
    Public Function fLI_MenkyoCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fLI_MenkyoCount"

            fLI_MenkyoCount = m_numMenkyoKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>�Ƌ����C���f�b�N�X���O�A�܂��́A�Ƌ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fLI_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLI_HospitalCD"

            If m_numMenkyoIdx = 0 OrElse m_numMenkyoKensu = 0 Then
                fLI_HospitalCD = ""
            Else
                fLI_HospitalCD = g_MenkyoInfo(m_numMenkyoIdx).strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>�Ƌ����C���f�b�N�X���O�A�܂��́A�Ƌ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fLI_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLI_StaffMngID"

            If m_numMenkyoIdx = 0 OrElse m_numMenkyoKensu = 0 Then
                fLI_StaffMngID = ""
            Else
                fLI_StaffMngID = g_MenkyoInfo(m_numMenkyoIdx).strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' �Ƌ��R�[�h���擾����
    ''' </summary>
    ''' <returns>�Ƌ��R�[�h</returns>
    ''' <remarks>�Ƌ����C���f�b�N�X���O�A�܂��́A�Ƌ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fLI_MenkyoCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLI_MenkyoCD"

            If m_numMenkyoIdx = 0 OrElse m_numMenkyoKensu = 0 Then
                fLI_MenkyoCD = ""
            Else
                fLI_MenkyoCD = g_MenkyoInfo(m_numMenkyoIdx).strCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ƌ����̂��擾����
    ''' </summary>
    ''' <returns>�Ƌ�����</returns>
    ''' <remarks>�Ƌ����C���f�b�N�X���O�A�܂��́A�Ƌ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fLI_MenkyoName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLI_MenkyoName"

            If m_numMenkyoIdx = 0 OrElse m_numMenkyoKensu = 0 Then
                fLI_MenkyoName = ""
            Else
                fLI_MenkyoName = g_MenkyoInfo(m_numMenkyoIdx).strName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ƌ��ԍ����擾����
    ''' </summary>
    ''' <returns>�Ƌ��ԍ�</returns>
    ''' <remarks>�Ƌ����C���f�b�N�X���O�A�܂��́A�Ƌ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fLI_MenkyoNo() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLI_MenkyoNo"

            If m_numMenkyoIdx = 0 OrElse m_numMenkyoKensu = 0 Then
                fLI_MenkyoNo = ""
            Else
                fLI_MenkyoNo = g_MenkyoInfo(m_numMenkyoIdx).strNo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '2012/10/25 fujisawa add st -------------------------------------------------------------------
    ''' <summary>
    ''' �s���{���R�[�h���擾����
    ''' </summary>
    ''' <returns>�s���{���R�[�h</returns>
    ''' <remarks>�Ƌ����C���f�b�N�X���O�A�܂��́A�Ƌ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fLI_JapanAreaCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLI_JapanAreCD"

            If m_numMenkyoIdx = 0 OrElse m_numMenkyoKensu = 0 Then
                fLI_JapanAreaCD = ""
            Else
                fLI_JapanAreaCD = g_MenkyoInfo(m_numMenkyoIdx).strJapanAreaCD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �s���{�����̂��擾����
    ''' </summary>
    ''' <returns>�s���{������</returns>
    ''' <remarks>�Ƌ����C���f�b�N�X���O�A�܂��́A�Ƌ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fLI_JapanAreaName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLI_JapanAreName"

            If m_numMenkyoIdx = 0 OrElse m_numMenkyoKensu = 0 Then
                fLI_JapanAreaName = ""
            Else
                fLI_JapanAreaName = g_MenkyoInfo(m_numMenkyoIdx).strJapanAreaName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function
    '2012/10/25 fujisawa add end ------------------------------------------------------------------
    ''' <summary>-
    ''' �擾�N�������擾����
    ''' </summary>
    ''' <returns>�擾�N����</returns>
    ''' <remarks>�Ƌ����C���f�b�N�X���O�A�܂��́A�Ƌ���񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fLI_GetDate() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fLI_GetDate"

            If m_numMenkyoIdx = 0 OrElse m_numMenkyoKensu = 0 Then
                fLI_GetDate = 0
            Else
                fLI_GetDate = g_MenkyoInfo(m_numMenkyoIdx).numGetDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���l���擾����
    ''' </summary>
    ''' <returns>���l</returns>
    ''' <remarks>�Ƌ����C���f�b�N�X���O�A�܂��́A�Ƌ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fLI_Bikou() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLI_Bikou"

            If m_numMenkyoIdx = 0 OrElse m_numMenkyoKensu = 0 Then
                fLI_Bikou = ""
            Else
                fLI_Bikou = g_MenkyoInfo(m_numMenkyoIdx).strBikou
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>�Ƌ����C���f�b�N�X���O�A�܂��́A�Ƌ���񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fLI_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fLI_FirstTime"

            If m_numMenkyoIdx = 0 OrElse m_numMenkyoKensu = 0 Then
                fLI_FirstTime = 0
            Else
                fLI_FirstTime = g_MenkyoInfo(m_numMenkyoIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>�Ƌ����C���f�b�N�X���O�A�܂��́A�Ƌ���񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fLI_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fLI_UpdTime"

            If m_numMenkyoIdx = 0 OrElse m_numMenkyoKensu = 0 Then
                fLI_UpdTime = 0
            Else
                fLI_UpdTime = g_MenkyoInfo(m_numMenkyoIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���i��񌏐����擾����
    ''' </summary>
    ''' <returns>���i��񌏐�</returns>
    ''' <remarks></remarks>
    Public Function fSH_ShikakuCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSH_ShikakuCount"

            fSH_ShikakuCount = m_numShikakuKensu

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>���i���C���f�b�N�X���O�A�܂��́A���i��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSH_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSH_HospitalCD"

            If m_numShikakuIdx = 0 OrElse m_numShikakuKensu = 0 Then
                fSH_HospitalCD = ""
            Else
                fSH_HospitalCD = g_ShikakuInfo(m_numShikakuIdx).strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>���i���C���f�b�N�X���O�A�܂��́A���i��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSH_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSH_StaffMngID"

            If m_numShikakuIdx = 0 OrElse m_numShikakuKensu = 0 Then
                fSH_StaffMngID = ""
            Else
                fSH_StaffMngID = g_ShikakuInfo(m_numShikakuIdx).strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' ���i�R�[�h���擾����
    ''' </summary>
    ''' <returns>���i�R�[�h</returns>
    ''' <remarks>���i���C���f�b�N�X���O�A�܂��́A���i��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSH_ShikakuCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSH_ShikakuCD"

            If m_numShikakuIdx = 0 OrElse m_numShikakuKensu = 0 Then
                fSH_ShikakuCD = ""
            Else
                fSH_ShikakuCD = g_ShikakuInfo(m_numShikakuIdx).strCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���i���̂��擾����
    ''' </summary>
    ''' <returns>���i����</returns>
    ''' <remarks>���i���C���f�b�N�X���O�A�܂��́A���i��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSH_ShikakuName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSH_ShikakuName"

            If m_numShikakuIdx = 0 OrElse m_numShikakuKensu = 0 Then
                fSH_ShikakuName = ""
            Else
                fSH_ShikakuName = g_ShikakuInfo(m_numShikakuIdx).strName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �擾�N�������擾����
    ''' </summary>
    ''' <returns>�擾�N����</returns>
    ''' <remarks>���i���C���f�b�N�X���O�A�܂��́A���i��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSH_GetDate() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSH_GetDate"

            If m_numShikakuIdx = 0 OrElse m_numShikakuKensu = 0 Then
                fSH_GetDate = 0
            Else
                fSH_GetDate = g_ShikakuInfo(m_numShikakuIdx).numGetDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �J�n�N�������擾����
    ''' </summary>
    ''' <returns>�J�n�N����</returns>
    ''' <remarks>���i���C���f�b�N�X���O�A�܂��́A���i��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSH_DateFrom() As Integer

        General.g_ErrorProc = "clsStaffIdo fSH_GetDate"
        Try
            If m_numShikakuIdx = 0 OrElse m_numShikakuKensu = 0 Then
                fSH_DateFrom = 0
            Else
                fSH_DateFrom = g_ShikakuInfo(m_numShikakuIdx).numDateFrom
            End If

        Catch
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �I���N�������擾����
    ''' </summary>
    ''' <returns>�I���N����</returns>
    ''' <remarks>���i���C���f�b�N�X���O�A�܂��́A���i��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSH_DateTo() As Integer

        General.g_ErrorProc = "clsStaffIdo fSH_DateTo"
        Try
            If m_numShikakuIdx = 0 OrElse m_numShikakuKensu = 0 Then
                fSH_DateTo = 0
            Else
                fSH_DateTo = g_ShikakuInfo(m_numShikakuIdx).numDateTo
            End If

        Catch
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���l���擾����
    ''' </summary>
    ''' <returns>���l</returns>
    ''' <remarks>���i���C���f�b�N�X���O�A�܂��́A���i��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSH_Bikou() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSH_Bikou"

            If m_numShikakuIdx = 0 OrElse m_numShikakuKensu = 0 Then
                fSH_Bikou = ""
            Else
                fSH_Bikou = g_ShikakuInfo(m_numShikakuIdx).strBikou
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>���i���C���f�b�N�X���O�A�܂��́A���i��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSH_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSH_FirstTime"

            If m_numShikakuIdx = 0 OrElse m_numShikakuKensu = 0 Then
                fSH_FirstTime = 0
            Else
                fSH_FirstTime = g_ShikakuInfo(m_numShikakuIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>���i���C���f�b�N�X���O�A�܂��́A���i��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSH_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSH_UpdTime"

            If m_numShikakuIdx = 0 OrElse m_numShikakuKensu = 0 Then
                fSH_UpdTime = 0
            Else
                fSH_UpdTime = g_ShikakuInfo(m_numShikakuIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ψ���񌏐����擾����
    ''' </summary>
    ''' <returns>�ψ���񌏐�</returns>
    ''' <remarks></remarks>
    Public Function fII_IinCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fII_IinCount"

            fII_IinCount = m_numIinKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�@�R�[�h</returns>
    ''' <remarks>�ψ����C���f�b�N�X���O�A�܂��́A�ψ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fII_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fII_HospitalCD"

            If m_numIinIdx = 0 OrElse m_numIinKensu = 0 Then
                fII_HospitalCD = ""
            Else
                fII_HospitalCD = g_IinInfo(m_numIinIdx).strHospitalCD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>�ψ����C���f�b�N�X���O�A�܂��́A�ψ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fII_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fII_StaffMngID"

            If m_numIinIdx = 0 OrElse m_numIinKensu = 0 Then
                fII_StaffMngID = ""
            Else
                fII_StaffMngID = g_IinInfo(m_numIinIdx).strStaffMngID
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' �ψ��R�[�h���擾����
    ''' </summary>
    ''' <returns>�ψ��R�[�h</returns>
    ''' <remarks>�ψ����C���f�b�N�X���O�A�܂��́A�ψ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fII_IinCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fII_IinCD"

            If m_numIinIdx = 0 OrElse m_numIinKensu = 0 Then
                fII_IinCD = ""
            Else
                fII_IinCD = g_IinInfo(m_numIinIdx).strCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ψ����̂��擾����
    ''' </summary>
    ''' <returns>�ψ�����</returns>
    ''' <remarks>�ψ����C���f�b�N�X���O�A�܂��́A�ψ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fII_IinName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fII_IinName"

            If m_numIinIdx = 0 OrElse m_numIinKensu = 0 Then
                fII_IinName = ""
            Else
                fII_IinName = g_IinInfo(m_numIinIdx).strName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �J�n�N�������擾����
    ''' </summary>
    ''' <returns>�J�n�N����</returns>
    ''' <remarks>�ψ����C���f�b�N�X���O�A�܂��́A�ψ���񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fII_DateFrom() As Integer

        General.g_ErrorProc = "clsStaffIdo fII_GetDate"
        Try
            If m_numIinIdx = 0 OrElse m_numIinKensu = 0 Then
                fII_DateFrom = (0)
            Else
                fII_DateFrom = g_IinInfo(m_numIinIdx).numDateFrom
            End If

        Catch
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �I���N�������擾����
    ''' </summary>
    ''' <returns>�I���N����</returns>
    ''' <remarks>�ψ����C���f�b�N�X���O�A�܂��́A�ψ���񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fII_DateTo() As Integer

        General.g_ErrorProc = "clsStaffIdo fII_GetDate"
        Try
            If m_numIinIdx = 0 OrElse m_numIinKensu = 0 Then
                fII_DateTo = 0
            Else
                fII_DateTo = g_IinInfo(m_numIinIdx).numDateTo
            End If
        Catch
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���l���擾����
    ''' </summary>
    ''' <returns>���l</returns>
    ''' <remarks>�ψ����C���f�b�N�X���O�A�܂��́A�ψ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fII_Bikou() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fII_Bikou"

            If m_numIinIdx = 0 OrElse m_numIinKensu = 0 Then
                fII_Bikou = ""
            Else
                fII_Bikou = g_IinInfo(m_numIinIdx).strBikou
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>�ψ����C���f�b�N�X���O�A�܂��́A�ψ���񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fII_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fII_FirstTime"

            If m_numIinIdx = 0 OrElse m_numIinKensu = 0 Then
                fII_FirstTime = 0
            Else
                fII_FirstTime = g_IinInfo(m_numIinIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>�ψ����C���f�b�N�X���O�A�܂��́A�ψ���񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fII_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fII_UpdTime"

            If m_numIinIdx = 0 OrElse m_numIinKensu = 0 Then
                fII_UpdTime = 0
            Else
                fII_UpdTime = g_IinInfo(m_numIinIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �ψ���E����
    ''' </summary>
    ''' <returns>�ψ���E����</returns>
    ''' <remarks>�ψ����C���f�b�N�X���O�A�܂��́A�ψ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fII_IinPostName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fII_IinPostName"

            If m_numIinIdx = 0 OrElse m_numIinKensu = 0 Then
                fII_IinPostName = ""
            Else
                fII_IinPostName = g_IinInfo(m_numIinIdx).strIinPostName
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �ψ���ECD
    ''' </summary>
    ''' <returns>�ψ���ECD</returns>
    ''' <remarks>�ψ����C���f�b�N�X���O�A�܂��́A�ψ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fII_IinPostCd() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fII_IinPostCd"

            If m_numIinIdx = 0 OrElse m_numIinKensu = 0 Then
                fII_IinPostCd = ""
            Else
                fII_IinPostCd = g_IinInfo(m_numIinIdx).strIinPostCd
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E����񌏐����擾����
    ''' </summary>
    ''' <returns>�E����񌏐�</returns>
    ''' <remarks></remarks>
    Public Function fJC_SyokurekiCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fJC_SyokurekiCount"

            fJC_SyokurekiCount = m_numSyokurekiKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>�E�����C���f�b�N�X���O�A�܂��́A�E����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fJC_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fJC_HospitalCD"

            If m_numSyokurekiIdx = 0 OrElse m_numSyokurekiKensu = 0 Then
                fJC_HospitalCD = ""
            Else
                fJC_HospitalCD = g_SyokurekiInfo(m_numSyokurekiIdx).strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>�E�����C���f�b�N�X���O�A�܂��́A�E����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fJC_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fJC_StaffMngID"

            If m_numSyokurekiIdx = 0 OrElse m_numSyokurekiKensu = 0 Then
                fJC_StaffMngID = ""
            Else
                fJC_StaffMngID = g_SyokurekiInfo(m_numSyokurekiIdx).strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ζ��n�R�[�h���擾����
    ''' </summary>
    ''' <returns>�Ζ��n�R�[�h</returns>
    ''' <remarks>�E�����C���f�b�N�X���O�A�܂��́A�E����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fJC_KinmuchiCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fJC_KinmuchiCD"

            If m_numSyokurekiIdx = 0 OrElse m_numSyokurekiKensu = 0 Then
                fJC_KinmuchiCD = ""
            Else
                fJC_KinmuchiCD = g_SyokurekiInfo(m_numSyokurekiIdx).strCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ζ��n���̂��擾����
    ''' </summary>
    ''' <returns>�Ζ��n����</returns>
    ''' <remarks>�E�����C���f�b�N�X���O�A�܂��́A�E����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fJC_KinmuchiName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fJC_KinmuchiName"

            If m_numSyokurekiIdx = 0 OrElse m_numSyokurekiKensu = 0 Then
                fJC_KinmuchiName = ""
            Else
                fJC_KinmuchiName = g_SyokurekiInfo(m_numSyokurekiIdx).strName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �J�n�N�������擾����
    ''' </summary>
    ''' <returns>�J�n�N����</returns>
    ''' <remarks>�E�����C���f�b�N�X���O�A�܂��́A�E����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fJC_DateFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fJC_GetDate"

            If m_numSyokurekiIdx = 0 OrElse m_numSyokurekiKensu = 0 Then
                fJC_DateFrom = 0
            Else
                fJC_DateFrom = g_SyokurekiInfo(m_numSyokurekiIdx).numDateFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �I���N�������擾����
    ''' </summary>
    ''' <returns>�I���N����</returns>
    ''' <remarks>�E�����C���f�b�N�X���O�A�܂��́A�E����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fJC_DateTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fJC_GetDate"

            If m_numSyokurekiIdx = 0 OrElse m_numSyokurekiKensu = 0 Then
                fJC_DateTo = 0
            Else
                fJC_DateTo = g_SyokurekiInfo(m_numSyokurekiIdx).numDateTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ζ��@�ւ��擾����
    ''' </summary>
    ''' <returns>�Ζ��@��</returns>
    ''' <remarks>�E�����C���f�b�N�X���O�A�܂��́A�E����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fJC_KinmuKikan() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fJC_KinmuKikan"

            If m_numSyokurekiIdx = 0 OrElse m_numSyokurekiKensu = 0 Then
                fJC_KinmuKikan = ""
            Else
                fJC_KinmuKikan = g_SyokurekiInfo(m_numSyokurekiIdx).strArea
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �o���f�ÉȂ��擾����
    ''' </summary>
    ''' <returns>�o���f�É�</returns>
    ''' <remarks>�E�����C���f�b�N�X���O�A�܂��́A�E����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fJC_ExpMedicalName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fJC_ExpMedicalName"

            If m_numSyokurekiIdx = 0 OrElse m_numSyokurekiKensu = 0 Then
                fJC_ExpMedicalName = ""
            Else
                fJC_ExpMedicalName = g_SyokurekiInfo(m_numSyokurekiIdx).strExpMedicalName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���l���擾����
    ''' </summary>
    ''' <returns>���l</returns>
    ''' <remarks>�E�����C���f�b�N�X���O�A�܂��́A�E����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fJC_Bikou() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fJC_Bikou"

            If m_numSyokurekiIdx = 0 OrElse m_numSyokurekiKensu = 0 Then
                fJC_Bikou = ""
            Else
                fJC_Bikou = g_SyokurekiInfo(m_numSyokurekiIdx).strBikou
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>�E�����C���f�b�N�X���O�A�܂��́A�E����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fJC_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fJC_FirstTime"

            If m_numSyokurekiIdx = 0 OrElse m_numSyokurekiKensu = 0 Then
                fJC_FirstTime = 0
            Else
                fJC_FirstTime = g_SyokurekiInfo(m_numSyokurekiIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>�E�����C���f�b�N�X���O�A�܂��́A�E����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fJC_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fJC_UpdTime"

            If m_numSyokurekiIdx = 0 OrElse m_numSyokurekiKensu = 0 Then
                fJC_UpdTime = 0
            Else
                fJC_UpdTime = g_SyokurekiInfo(m_numSyokurekiIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��ʊw����񌏐����擾����
    ''' </summary>
    ''' <returns>��ʊw����񌏐�</returns>
    ''' <remarks></remarks>
    Public Function fGS_IppanGakurekiCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fGS_IppanGakurekiCount"

            fGS_IppanGakurekiCount = m_numIppanGakurekiKensu

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>��ʊw�����C���f�b�N�X���O�A�܂��́A��ʊw����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fGS_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGS_HospitalCD"

            If m_numIppanGakurekiIdx = 0 OrElse m_numIppanGakurekiKensu = 0 Then
                fGS_HospitalCD = ""
            Else
                fGS_HospitalCD = g_IppanGakurekiInfo(m_numIppanGakurekiIdx).strHospitalCD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>��ʊw�����C���f�b�N�X���O�A�܂��́A��ʊw����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fGS_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGS_StaffMngID"

            If m_numIppanGakurekiIdx = 0 OrElse m_numIppanGakurekiKensu = 0 Then
                fGS_StaffMngID = ""
            Else
                fGS_StaffMngID = g_IppanGakurekiInfo(m_numIppanGakurekiIdx).strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �敪���擾����
    ''' </summary>
    ''' <returns>�敪</returns>
    ''' <remarks>��ʊw�����C���f�b�N�X���O�A�܂��́A��ʊw����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fGS_Kbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGS_Kbn"

            If m_numIppanGakurekiIdx = 0 OrElse m_numIppanGakurekiKensu = 0 Then
                fGS_Kbn = ""
            Else
                fGS_Kbn = g_IppanGakurekiInfo(m_numIppanGakurekiIdx).strKbn
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �敪���̂��擾����
    ''' </summary>
    ''' <returns>�敪����</returns>
    ''' <remarks>��ʊw�����C���f�b�N�X���O�A�܂��́A��ʊw����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fGS_KbnName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGS_KbnName"

            If m_numIppanGakurekiIdx = 0 OrElse m_numIppanGakurekiKensu = 0 Then
                fGS_KbnName = ""
            Else
                fGS_KbnName = g_IppanGakurekiInfo(m_numIppanGakurekiIdx).strKbnName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �n��R�[�h���擾����
    ''' </summary>
    ''' <returns>�n��R�[�h</returns>
    ''' <remarks>��ʊw�����C���f�b�N�X���O�A�܂��́A��ʊw����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fGS_ChiikiCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGS_ChiikiCD"

            If m_numIppanGakurekiIdx = 0 OrElse m_numIppanGakurekiKensu = 0 Then
                fGS_ChiikiCD = ""
            Else
                fGS_ChiikiCD = g_IppanGakurekiInfo(m_numIppanGakurekiIdx).strChiikiCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �n�於�̂��擾����
    ''' </summary>
    ''' <returns>�n�於��</returns>
    ''' <remarks>��ʊw�����C���f�b�N�X���O�A�܂��́A��ʊw����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fGS_ChiikiName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGS_ChiikiName"

            If m_numIppanGakurekiIdx = 0 OrElse m_numIppanGakurekiKensu = 0 Then
                fGS_ChiikiName = ""
            Else
                fGS_ChiikiName = g_IppanGakurekiInfo(m_numIppanGakurekiIdx).strChiikiName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�w���敪���擾����
    ''' </summary>
    ''' <returns>�ŏI�w���敪</returns>
    ''' <remarks>��ʊw�����C���f�b�N�X���O�A�܂��́A��ʊw����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fGS_LastKbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGS_LastKbn"

            If m_numIppanGakurekiIdx = 0 OrElse m_numIppanGakurekiKensu = 0 Then
                fGS_LastKbn = ""
            Else
                fGS_LastKbn = g_IppanGakurekiInfo(m_numIppanGakurekiIdx).strLastKbn
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���ƔN�������擾����
    ''' </summary>
    ''' <returns>���ƔN����</returns>
    ''' <remarks>��ʊw�����C���f�b�N�X���O�A�܂��́A��ʊw����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fGS_LastDate() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fGS_LastDate"

            If m_numIppanGakurekiIdx = 0 OrElse m_numIppanGakurekiKensu = 0 Then
                fGS_LastDate = 0
            Else
                fGS_LastDate = g_IppanGakurekiInfo(m_numIppanGakurekiIdx).numDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' �w�Z�R�[�h���擾����
    ''' </summary>
    ''' <returns>�w�Z�R�[�h</returns>
    ''' <remarks>��ʊw�����C���f�b�N�X���O�A�܂��́A��ʊw����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fGS_SchoolCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGS_SchoolCD"

            If m_numIppanGakurekiIdx = 0 OrElse m_numIppanGakurekiKensu = 0 Then
                fGS_SchoolCD = ""
            Else
                fGS_SchoolCD = g_IppanGakurekiInfo(m_numIppanGakurekiIdx).strSchoolCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �w�Z�����擾����
    ''' </summary>
    ''' <returns>�w�Z��</returns>
    ''' <remarks>��ʊw�����C���f�b�N�X���O�A�܂��́A��ʊw����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fGS_SchoolName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGS_SchoolName"

            If m_numIppanGakurekiIdx = 0 OrElse m_numIppanGakurekiKensu = 0 Then
                fGS_SchoolName = ""
            Else
                fGS_SchoolName = g_IppanGakurekiInfo(m_numIppanGakurekiIdx).strSchoolName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���l���擾����
    ''' </summary>
    ''' <returns>���l</returns>
    ''' <remarks>��ʊw�����C���f�b�N�X���O�A�܂��́A��ʊw����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fGS_Bikou() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGS_Bikou"

            If m_numIppanGakurekiIdx = 0 OrElse m_numIppanGakurekiKensu = 0 Then
                fGS_Bikou = ""
            Else
                fGS_Bikou = g_IppanGakurekiInfo(m_numIppanGakurekiIdx).strBikou
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>��ʊw�����C���f�b�N�X���O�A�܂��́A��ʊw����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fGS_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fGS_FirstTime"

            If m_numIppanGakurekiIdx = 0 OrElse m_numIppanGakurekiKensu = 0 Then
                fGS_FirstTime = 0
            Else
                fGS_FirstTime = g_IppanGakurekiInfo(m_numIppanGakurekiIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>��ʊw�����C���f�b�N�X���O�A�܂��́A��ʊw����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fGS_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fGS_UpdTime"

            If m_numIppanGakurekiIdx = 0 OrElse m_numIppanGakurekiKensu = 0 Then
                fGS_UpdTime = 0
            Else
                fGS_UpdTime = g_IppanGakurekiInfo(m_numIppanGakurekiIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���w����񌏐����擾����
    ''' </summary>
    ''' <returns>���w����񌏐�</returns>
    ''' <remarks></remarks>
    Public Function fSS_SenmonGakurekiCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSS_SenmonGakurekiCount"

            fSS_SenmonGakurekiCount = m_numSenmonGakurekiKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>���w�����C���f�b�N�X���O�A�܂��́A���w����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSS_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSS_HospitalCD"

            If m_numSenmonGakurekiIdx = 0 OrElse m_numSenmonGakurekiKensu = 0 Then
                fSS_HospitalCD = ""
            Else
                fSS_HospitalCD = g_SenmonGakurekiInfo(m_numSenmonGakurekiIdx).strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>���w�����C���f�b�N�X���O�A�܂��́A���w����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSS_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSS_StaffMngID"

            If m_numSenmonGakurekiIdx = 0 OrElse m_numSenmonGakurekiKensu = 0 Then
                fSS_StaffMngID = ""
            Else
                fSS_StaffMngID = g_SenmonGakurekiInfo(m_numSenmonGakurekiIdx).strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �敪���擾����
    ''' </summary>
    ''' <returns>�敪</returns>
    ''' <remarks>���w�����C���f�b�N�X���O�A�܂��́A���w����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSS_Kbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSS_Kbn"

            If m_numSenmonGakurekiIdx = 0 OrElse m_numSenmonGakurekiKensu = 0 Then
                fSS_Kbn = ""
            Else
                fSS_Kbn = g_SenmonGakurekiInfo(m_numSenmonGakurekiIdx).strKbn
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �敪���̂��擾����
    ''' </summary>
    ''' <returns>�敪����</returns>
    ''' <remarks>���w�����C���f�b�N�X���O�A�܂��́A���w����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSS_KbnName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSS_KbnName"

            If m_numSenmonGakurekiIdx = 0 OrElse m_numSenmonGakurekiKensu = 0 Then
                fSS_KbnName = ""
            Else
                fSS_KbnName = g_SenmonGakurekiInfo(m_numSenmonGakurekiIdx).strKbnName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �n��R�[�h���擾����
    ''' </summary>
    ''' <returns>�n��R�[�h</returns>
    ''' <remarks>���w�����C���f�b�N�X���O�A�܂��́A���w����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSS_ChiikiCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSS_ChiikiCD"

            If m_numSenmonGakurekiIdx = 0 OrElse m_numSenmonGakurekiKensu = 0 Then
                fSS_ChiikiCD = ""
            Else
                fSS_ChiikiCD = g_SenmonGakurekiInfo(m_numSenmonGakurekiIdx).strChiikiCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �n�於�̂��擾����
    ''' </summary>
    ''' <returns>�n�於��</returns>
    ''' <remarks>���w�����C���f�b�N�X���O�A�܂��́A���w����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSS_ChiikiName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSS_ChiikiName"

            If m_numSenmonGakurekiIdx = 0 OrElse m_numSenmonGakurekiKensu = 0 Then
                fSS_ChiikiName = ""
            Else
                fSS_ChiikiName = g_SenmonGakurekiInfo(m_numSenmonGakurekiIdx).strChiikiName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�w���敪���擾����
    ''' </summary>
    ''' <returns>�ŏI�w���敪</returns>
    ''' <remarks>���w�����C���f�b�N�X���O�A�܂��́A���w����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSS_LastKbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSS_LastKbn"

            If m_numSenmonGakurekiIdx = 0 OrElse m_numSenmonGakurekiKensu = 0 Then
                fSS_LastKbn = ""
            Else
                fSS_LastKbn = g_SenmonGakurekiInfo(m_numSenmonGakurekiIdx).strLastKbn
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���ƔN�������擾����
    ''' </summary>
    ''' <returns>���ƔN����</returns>
    ''' <remarks>���w�����C���f�b�N�X���O�A�܂��́A���w����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSS_LastDate() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSS_LastDate"

            If m_numSenmonGakurekiIdx = 0 OrElse m_numSenmonGakurekiKensu = 0 Then
                fSS_LastDate = 0
            Else
                fSS_LastDate = g_SenmonGakurekiInfo(m_numSenmonGakurekiIdx).numDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' �w�Z�R�[�h���擾����
    ''' </summary>
    ''' <returns>�w�Z�R�[�h</returns>
    ''' <remarks>���w�����C���f�b�N�X���O�A�܂��́A���w����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSS_SchoolCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSS_SchoolCD"

            If m_numSenmonGakurekiIdx = 0 OrElse m_numSenmonGakurekiKensu = 0 Then
                fSS_SchoolCD = ""
            Else
                fSS_SchoolCD = g_SenmonGakurekiInfo(m_numSenmonGakurekiIdx).strSchoolCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �w�Z�����擾����
    ''' </summary>
    ''' <returns>�w�Z��</returns>
    ''' <remarks>���w�����C���f�b�N�X���O�A�܂��́A���w����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSS_SchoolName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSS_SchoolName"

            If m_numSenmonGakurekiIdx = 0 OrElse m_numSenmonGakurekiKensu = 0 Then
                fSS_SchoolName = ""
            Else
                fSS_SchoolName = g_SenmonGakurekiInfo(m_numSenmonGakurekiIdx).strSchoolName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���l���擾����
    ''' </summary>
    ''' <returns>���l</returns>
    ''' <remarks>���w�����C���f�b�N�X���O�A�܂��́A���w����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSS_Bikou() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSS_Bikou"

            If m_numSenmonGakurekiIdx = 0 OrElse m_numSenmonGakurekiKensu = 0 Then
                fSS_Bikou = ""
            Else
                fSS_Bikou = g_SenmonGakurekiInfo(m_numSenmonGakurekiIdx).strBikou
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>���w�����C���f�b�N�X���O�A�܂��́A���w����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSS_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSS_FirstTime"

            If m_numSenmonGakurekiIdx = 0 OrElse m_numSenmonGakurekiKensu = 0 Then
                fSS_FirstTime = 0
            Else
                fSS_FirstTime = g_SenmonGakurekiInfo(m_numSenmonGakurekiIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>���w�����C���f�b�N�X���O�A�܂��́A���w����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSS_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSS_UpdTime"

            If m_numSenmonGakurekiIdx = 0 OrElse m_numSenmonGakurekiKensu = 0 Then
                fSS_UpdTime = 0
            Else
                fSS_UpdTime = g_SenmonGakurekiInfo(m_numSenmonGakurekiIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���x��񌏐����擾����
    ''' </summary>
    ''' <returns>���x��񌏐�</returns>
    ''' <remarks></remarks>
    Public Function fLL_ChoukyuCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fLL_ChoukyuCount"

            fLL_ChoukyuCount = m_numChoukyuKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>���x���C���f�b�N�X���O�A�܂��́A���x��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fLL_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLL_HospitalCD"

            If m_numChoukyuIdx = 0 OrElse m_numChoukyuKensu = 0 Then
                fLL_HospitalCD = ""
            Else
                fLL_HospitalCD = g_ChoukyuInfo(m_numChoukyuIdx).strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>���x���C���f�b�N�X���O�A�܂��́A���x��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fLL_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLL_StaffMngID"

            If m_numChoukyuIdx = 0 OrElse m_numChoukyuKensu = 0 Then
                fLL_StaffMngID = ""
            Else
                fLL_StaffMngID = g_ChoukyuInfo(m_numChoukyuIdx).strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �R�[�h���擾����
    ''' </summary>
    ''' <returns>�R�[�h</returns>
    ''' <remarks>���x���C���f�b�N�X���O�A�܂��́A���x��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fLL_CD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLL_CD"

            If m_numChoukyuIdx = 0 OrElse m_numChoukyuKensu = 0 Then
                fLL_CD = ""
            Else
                fLL_CD = g_ChoukyuInfo(m_numChoukyuIdx).strCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���̂��擾����
    ''' </summary>
    ''' <returns>����</returns>
    ''' <remarks>���x���C���f�b�N�X���O�A�܂��́A���x��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fLL_Name() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLL_Name"

            If m_numChoukyuIdx = 0 OrElse m_numChoukyuKensu = 0 Then
                fLL_Name = ""
            Else
                fLL_Name = g_ChoukyuInfo(m_numChoukyuIdx).strName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    '2018/10/02 Darren ADD START
    Public Function fLL_SecName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLL_SecName"

            If m_numChoukyuIdx = 0 OrElse m_numChoukyuKensu = 0 Then
                fLL_SecName = ""
            Else
                fLL_SecName = g_ChoukyuInfo(m_numChoukyuIdx).strSecName
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    '2018/10/02 Darren ADD END

    ''' <summary>
    ''' �J�n�N�������擾����
    ''' </summary>
    ''' <returns>�J�n�N����</returns>
    ''' <remarks>���x���C���f�b�N�X���O�A�܂��́A���x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fLL_DateFrom() As Integer

        General.g_ErrorProc = "clsStaffIdo fLL_DateFrom"
        Try
            If m_numChoukyuIdx = 0 OrElse m_numChoukyuKensu = 0 Then
                fLL_DateFrom = 0
            Else
                fLL_DateFrom = g_ChoukyuInfo(m_numChoukyuIdx).numDateFrom
            End If

        Catch
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �I���N�������擾����
    ''' </summary>
    ''' <returns>�I���N����</returns>
    ''' <remarks>���x���C���f�b�N�X���O�A�܂��́A���x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fLL_DateTo() As Integer

        General.g_ErrorProc = "clsStaffIdo fLL_DateTo"
        Try
            If m_numChoukyuIdx = 0 OrElse m_numChoukyuKensu = 0 Then
                fLL_DateTo = 0
            Else
                fLL_DateTo = g_ChoukyuInfo(m_numChoukyuIdx).numDateTo
            End If

        Catch
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���l���擾����
    ''' </summary>
    ''' <returns>���l</returns>
    ''' <remarks>���x���C���f�b�N�X���O�A�܂��́A���x��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fLL_Bikou() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLL_Bikou"

            If m_numChoukyuIdx = 0 OrElse m_numChoukyuKensu = 0 Then
                fLL_Bikou = ""
            Else
                fLL_Bikou = g_ChoukyuInfo(m_numChoukyuIdx).strBikou
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    '2018/08/24 T.K add st --------------------------------
    ''' <summary>
    ''' �T�J�����Ԃ��擾����
    ''' </summary>
    ''' <returns>�T�J������</returns>
    ''' <remarks>���x���C���f�b�N�X���O�A�܂��́A���x��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fLL_WeeklyTime() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fLL_WeeklyTime"

            If m_numChoukyuIdx = 0 OrElse m_numChoukyuKensu = 0 Then
                fLL_WeeklyTime = ""
            Else
                fLL_WeeklyTime = g_ChoukyuInfo(m_numChoukyuIdx).numWeeklyTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '2018/08/24 T.K add ed --------------------------------

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>���x���C���f�b�N�X���O�A�܂��́A���x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fLL_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fLL_FirstTime"

            If m_numChoukyuIdx = 0 OrElse m_numChoukyuKensu = 0 Then
                fLL_FirstTime = 0
            Else
                fLL_FirstTime = g_ChoukyuInfo(m_numChoukyuIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>���x���C���f�b�N�X���O�A�܂��́A���x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fLL_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fLL_UpdTime"

            If m_numChoukyuIdx = 0 OrElse m_numChoukyuKensu = 0 Then
                fLL_UpdTime = 0
            Else
                fLL_UpdTime = g_ChoukyuInfo(m_numChoukyuIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Y�x��񌏐����擾����
    ''' </summary>
    ''' <returns>�Y�x��񌏐�</returns>
    ''' <remarks></remarks>
    Public Function fSK_SankyuCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_SankyuCount"

            fSK_SankyuCount = m_numSankyuKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSK_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_HospitalCD"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_HospitalCD = ""
            Else
                fSK_HospitalCD = g_SankyuInfo(m_numSankyuIdx).strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSK_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_StaffMngID"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_StaffMngID = ""
            Else
                fSK_StaffMngID = g_SankyuInfo(m_numSankyuIdx).strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' SEQ���擾����
    ''' </summary>
    ''' <returns>SEQ</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_SEQ() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_SEQ"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_SEQ = 0
            Else
                fSK_SEQ = g_SankyuInfo(m_numSankyuIdx).numSEQ
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �\��N�������擾����
    ''' </summary>
    ''' <returns>�\��N����</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_PlanDate() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_PlanDate"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_PlanDate = 0
            Else
                fSK_PlanDate = g_SankyuInfo(m_numSankyuIdx).numPlanDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �o�ً敪���擾����
    ''' </summary>
    ''' <returns>�o�ً敪</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSK_TwinFlg() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_TwinFlg"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_TwinFlg = ""
            Else
                fSK_TwinFlg = g_SankyuInfo(m_numSankyuIdx).strTwinFlg
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �o�Y�N�������擾����
    ''' </summary>
    ''' <returns>�o�Y�N����</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_BirthDate() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_BirthDate"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_BirthDate = 0
            Else
                fSK_BirthDate = g_SankyuInfo(m_numSankyuIdx).numBirthDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �\��Y�O���From���擾����
    ''' </summary>
    ''' <returns>�\��Y�O���From</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_PlanSanzenYamenFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_PlanSanzenYamenFrom"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_PlanSanzenYamenFrom = 0
            Else
                fSK_PlanSanzenYamenFrom = g_SankyuInfo(m_numSankyuIdx).numPlanSanzenYamenFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �\��Y�O���To���擾����
    ''' </summary>
    ''' <returns>�\��Y�O���To</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_PlanSanzenYamenTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_PlanSanzenYamenTo"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_PlanSanzenYamenTo = 0
            Else
                fSK_PlanSanzenYamenTo = g_SankyuInfo(m_numSankyuIdx).numPlanSanzenYamenTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �\��Y�O�x��From���擾����
    ''' </summary>
    ''' <returns>�\��Y�O�x��From</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_PlanSanzenHolFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_PlanSanzenHolFrom"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_PlanSanzenHolFrom = 0
            Else
                fSK_PlanSanzenHolFrom = g_SankyuInfo(m_numSankyuIdx).numPlanSanzenHolFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �\��Y�O�x��To���擾����
    ''' </summary>
    ''' <returns>�\��Y�O�x��To</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_PlanSanzenHolTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_PlanSanzenHolTo"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_PlanSanzenHolTo = 0
            Else
                fSK_PlanSanzenHolTo = g_SankyuInfo(m_numSankyuIdx).numPlanSanzenHolTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �\��Y��x��From���擾����
    ''' </summary>
    ''' <returns>�\��Y��x��From</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_PlanSangoHolFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_PlanSangoHolFrom"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_PlanSangoHolFrom = 0
            Else
                fSK_PlanSangoHolFrom = g_SankyuInfo(m_numSankyuIdx).numPlanSangoHolFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �\��Y��x��To���擾����
    ''' </summary>
    ''' <returns>�\��Y��x��To</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_PlanSangoHolTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_PlanSangoHolTo"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_PlanSangoHolTo = 0
            Else
                fSK_PlanSangoHolTo = g_SankyuInfo(m_numSankyuIdx).numPlanSangoHolTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    '
    ''' <summary>
    ''' �\��玙�x��From���擾����
    ''' </summary>
    ''' <returns>�\��玙�x��From</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_PlanIkujiHolFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_PlanIkujiHolFrom"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_PlanIkujiHolFrom = 0
            Else
                fSK_PlanIkujiHolFrom = g_SankyuInfo(m_numSankyuIdx).numPlanIkujiHolFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �\��玙�x��To���擾����
    ''' </summary>
    ''' <returns>�\��玙�x��To</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_PlanIkujiHolTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_PlanIkujiHolTo"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_PlanIkujiHolTo = 0
            Else
                fSK_PlanIkujiHolTo = g_SankyuInfo(m_numSankyuIdx).numPlanIkujiHolTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '

    ''' <summary>
    ''' �m��Y�O���From���擾����
    ''' </summary>
    ''' <returns>�m��Y�O���From</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_FixedSanzenYamenFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_FixedSanzenYamenFrom"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_FixedSanzenYamenFrom = 0
            Else
                fSK_FixedSanzenYamenFrom = g_SankyuInfo(m_numSankyuIdx).numFixedSanzenYamenFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �m��Y�O���To���擾����
    ''' </summary>
    ''' <returns>�m��Y�O���To</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_FixedSanzenYamenTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_FixedSanzenYamenTo"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_FixedSanzenYamenTo = 0
            Else
                fSK_FixedSanzenYamenTo = g_SankyuInfo(m_numSankyuIdx).numFixedSanzenYamenTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �m��Y�O�x��From���擾����
    ''' </summary>
    ''' <returns>�m��Y�O�x��From</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_FixedSanzenHolFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_FixedSanzenHolFrom"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_FixedSanzenHolFrom = 0
            Else
                fSK_FixedSanzenHolFrom = g_SankyuInfo(m_numSankyuIdx).numFixedSanzenHolFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �m��Y�O�x��To���擾����
    ''' </summary>
    ''' <returns>�m��Y�O�x��To</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_FixedSanzenHolTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_FixedSanzenHolTo"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_FixedSanzenHolTo = 0
            Else
                fSK_FixedSanzenHolTo = g_SankyuInfo(m_numSankyuIdx).numFixedSanzenHolTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �m��Y��x��From���擾����
    ''' </summary>
    ''' <returns>�m��Y��x��From</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_FixedSangoHolFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_FixedSangoHolFrom"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_FixedSangoHolFrom = 0
            Else
                fSK_FixedSangoHolFrom = g_SankyuInfo(m_numSankyuIdx).numFixedSangoHolFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �m��Y��x��To���擾����
    ''' </summary>
    ''' <returns>�m��Y��x��To</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_FixedSangoHolTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_FixedSangoHolTo"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_FixedSangoHolTo = 0
            Else
                fSK_FixedSangoHolTo = g_SankyuInfo(m_numSankyuIdx).numFixedSangoHolTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    '
    ''' <summary>
    ''' �m��玙�x��From���擾����
    ''' </summary>
    ''' <returns>�m��玙�x��From</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_FixedIkujiHolFrom() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_FixedIkujiHolFrom"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_FixedIkujiHolFrom = 0
            Else
                fSK_FixedIkujiHolFrom = g_SankyuInfo(m_numSankyuIdx).numFixedIkujiHolFrom
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �m��玙�x��To���擾����
    ''' </summary>
    ''' <returns>�m��玙�x��To</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_FixedIkujiHolTo() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_FixedIkujiHolTo"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_FixedIkujiHolTo = 0
            Else
                fSK_FixedIkujiHolTo = g_SankyuInfo(m_numSankyuIdx).numFixedIkujiHolTo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�����擾����
    ''' </summary>
    ''' <returns>����o�^��</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_FirstTime"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_FirstTime = 0
            Else
                fSK_FirstTime = g_SankyuInfo(m_numSankyuIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_UpdTime"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_UpdTime = 0
            Else
                fSK_UpdTime = g_SankyuInfo(m_numSankyuIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' UniqueSeqNO���擾����
    ''' </summary>
    ''' <returns>UniqueSeqNO</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_UniqueSeqNO() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_UniqueSeqNO"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_UniqueSeqNO = "0"
            Else
                fSK_UniqueSeqNO = g_SankyuInfo(m_numSankyuIdx).strUniqueSeqNO
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���F�σt���O���擾����
    ''' </summary>
    ''' <returns>���F�σt���O</returns>
    ''' <remarks>�Y�x���C���f�b�N�X���O�A�܂��́A�Y�x��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSK_ApproveFlg() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSK_ApproveFlg"

            If m_numSankyuIdx = 0 OrElse m_numSankyuKensu = 0 Then
                fSK_ApproveFlg = "0"
            Else
                fSK_ApproveFlg = g_SankyuInfo(m_numSankyuIdx).strApproveFlg
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Y�x��񌏐����擾����
    ''' </summary>
    ''' <returns>�Y�x��񌏐�</returns>
    ''' <remarks></remarks>
    Public Function fSO_KyoukaiCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSO_KyoukaiCount"

            fSO_KyoukaiCount = m_numKyoukaiKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>������C���f�b�N�X���O�A�܂��́A�����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSO_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSO_HospitalCD"

            If m_numKyoukaiIdx = 0 OrElse m_numKyoukaiKensu = 0 Then
                fSO_HospitalCD = ""
            Else
                fSO_HospitalCD = g_KyoukaiInfo(m_numKyoukaiIdx).strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>������C���f�b�N�X���O�A�܂��́A�����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSO_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSO_StaffMngID"

            If m_numKyoukaiIdx = 0 OrElse m_numKyoukaiKensu = 0 Then
                fSO_StaffMngID = ""
            Else
                fSO_StaffMngID = g_KyoukaiInfo(m_numKyoukaiIdx).strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' ����R�[�h���擾����
    ''' </summary>
    ''' <returns>����R�[�h</returns>
    ''' <remarks>������C���f�b�N�X���O�A�܂��́A�����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSO_CD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSO_CD"

            If m_numKyoukaiIdx = 0 OrElse m_numKyoukaiKensu = 0 Then
                fSO_CD = ""
            Else
                fSO_CD = g_KyoukaiInfo(m_numKyoukaiIdx).strCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' ����̂��擾����
    ''' </summary>
    ''' <returns>�����</returns>
    ''' <remarks>������C���f�b�N�X���O�A�܂��́A�����񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSO_Name() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSO_Name"

            If m_numKyoukaiIdx = 0 OrElse m_numKyoukaiKensu = 0 Then
                fSO_Name = ""
            Else
                fSO_Name = g_KyoukaiInfo(m_numKyoukaiIdx).strName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' ����N�������擾����
    ''' </summary>
    ''' <returns>����N����</returns>
    ''' <remarks>������C���f�b�N�X���O�A�܂��́A�����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSO_Date() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSO_Date"

            If m_numKyoukaiIdx = 0 OrElse m_numKyoukaiKensu = 0 Then
                fSO_Date = 0
            Else
                fSO_Date = g_KyoukaiInfo(m_numKyoukaiIdx).numGetDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �މ�N�������擾����
    ''' </summary>
    ''' <returns>�މ�N����</returns>
    ''' <remarks>������C���f�b�N�X���O�A�܂��́A�����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSO_WithDrawDate() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSO_WithDrawDate"

            If m_numKyoukaiIdx = 0 OrElse m_numKyoukaiKensu = 0 Then
                fSO_WithDrawDate = 0
            Else
                fSO_WithDrawDate = g_KyoukaiInfo(m_numKyoukaiIdx).numEndDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����ԍ����擾����
    ''' </summary>
    ''' <returns>����ԍ�</returns>
    ''' <remarks>������C���f�b�N�X���O�A�܂��́A�����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSO_No() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSO_No"

            If m_numKyoukaiIdx = 0 OrElse m_numKyoukaiKensu = 0 Then
                fSO_No = "0"
            Else
                fSO_No = g_KyoukaiInfo(m_numKyoukaiIdx).strNo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>������C���f�b�N�X���O�A�܂��́A�����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSO_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSO_FirstTime"

            If m_numKyoukaiIdx = 0 OrElse m_numKyoukaiKensu = 0 Then
                fSO_FirstTime = 0
            Else
                fSO_FirstTime = g_KyoukaiInfo(m_numKyoukaiIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>������C���f�b�N�X���O�A�܂��́A�����񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSO_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fSO_UpdTime"

            If m_numKyoukaiIdx = 0 OrElse m_numKyoukaiKensu = 0 Then
                fSO_UpdTime = 0
            Else
                fSO_UpdTime = g_KyoukaiInfo(m_numKyoukaiIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ƒ���񌏐����擾����
    ''' </summary>
    ''' <returns>�Ƒ���񌏐�</returns>
    ''' <remarks></remarks>
    Public Function fKY_KazokuCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fKY_KazokuCount"

            fKY_KazokuCount = m_numKazokuKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�a�@�R�[�h</returns>
    ''' <remarks>�Ƒ����C���f�b�N�X���O�A�܂��́A�Ƒ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fKY_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKY_HospitalCD"

            If m_numKazokuIdx = 0 OrElse m_numKazokuKensu = 0 Then
                fKY_HospitalCD = ""
            Else
                fKY_HospitalCD = g_KazokuInfo(m_numKazokuIdx).strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks>�Ƒ����C���f�b�N�X���O�A�܂��́A�Ƒ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fKY_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKY_StaffMngID"

            If m_numKazokuIdx = 0 OrElse m_numKazokuKensu = 0 Then
                fKY_StaffMngID = ""
            Else
                fKY_StaffMngID = g_KazokuInfo(m_numKazokuIdx).strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '

    ''' <summary>
    ''' �Ƒ��������擾����
    ''' </summary>
    ''' <returns>�Ƒ�����</returns>
    ''' <remarks>�Ƒ����C���f�b�N�X���O�A�܂��́A�Ƒ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fKY_Name() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKY_Name"

            If m_numKazokuIdx = 0 OrElse m_numKazokuKensu = 0 Then
                fKY_Name = ""
            Else
                fKY_Name = g_KazokuInfo(m_numKazokuIdx).strName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' SEQ���擾����
    ''' </summary>
    ''' <returns>SEQ</returns>
    ''' <remarks>�Ƒ����C���f�b�N�X���O�A�܂��́A�Ƒ���񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fKY_SEQ() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKY_SEQ"

            If m_numKazokuIdx = 0 OrElse m_numKazokuKensu = 0 Then
                fKY_SEQ = (0)
            Else
                fKY_SEQ = (g_KazokuInfo(m_numKazokuIdx).numSEQ)
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���N�������擾����
    ''' </summary>
    ''' <returns>���N����</returns>
    ''' <remarks>�Ƒ����C���f�b�N�X���O�A�܂��́A�Ƒ���񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fKY_BirthDay() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fKY_BirthDay"

            If m_numKazokuIdx = 0 OrElse m_numKazokuKensu = 0 Then
                fKY_BirthDay = 0
            Else
                fKY_BirthDay = g_KazokuInfo(m_numKazokuIdx).numDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    '
    ''' <summary>
    ''' �����R�[�h���擾����
    ''' </summary>
    ''' <returns>�����R�[�h</returns>
    ''' <remarks>�Ƒ����C���f�b�N�X���O�A�܂��́A�Ƒ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fKY_TsudukiCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKY_TsudukiCD"

            If m_numKazokuIdx = 0 OrElse m_numKazokuKensu = 0 Then
                fKY_TsudukiCD = ""
            Else
                fKY_TsudukiCD = g_KazokuInfo(m_numKazokuIdx).strTsudukiCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �������̂��擾����
    ''' </summary>
    ''' <returns>��������</returns>
    ''' <remarks>�Ƒ����C���f�b�N�X���O�A�܂��́A�Ƒ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fKY_TsudukiName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKY_TsudukiName"

            If m_numKazokuIdx = 0 OrElse m_numKazokuKensu = 0 Then
                fKY_TsudukiName = ""
            Else
                fKY_TsudukiName = g_KazokuInfo(m_numKazokuIdx).strTsudukiName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �����敪���擾����
    ''' </summary>
    ''' <returns>�����敪</returns>
    ''' <remarks>�Ƒ����C���f�b�N�X���O�A�܂��́A�Ƒ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fKY_DoukyoKbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKY_DoukyoKbn"

            If m_numKazokuIdx = 0 OrElse m_numKazokuKensu = 0 Then
                fKY_DoukyoKbn = ""
            Else
                fKY_DoukyoKbn = g_KazokuInfo(m_numKazokuIdx).strDoukyoKbn
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �}�{�敪���擾����
    ''' </summary>
    ''' <returns>�}�{�敪</returns>
    ''' <remarks>�Ƒ����C���f�b�N�X���O�A�܂��́A�Ƒ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fKY_FuyouKbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKY_FuyouKbn"

            If m_numKazokuIdx = 0 OrElse m_numKazokuKensu = 0 Then
                fKY_FuyouKbn = ""
            Else
                fKY_FuyouKbn = g_KazokuInfo(m_numKazokuIdx).strFuyouKbn
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �����敪���擾����
    ''' </summary>
    ''' <returns>�����敪</returns>
    ''' <remarks>�Ƒ����C���f�b�N�X���O�A�܂��́A�Ƒ���񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fKY_SeizonKbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKY_SeizonKbn"

            If m_numKazokuIdx = 0 OrElse m_numKazokuKensu = 0 Then
                fKY_SeizonKbn = ""
            Else
                fKY_SeizonKbn = g_KazokuInfo(m_numKazokuIdx).strSeizonKbn
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks>�Ƒ����C���f�b�N�X���O�A�܂��́A�Ƒ���񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fKY_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fKY_FirstTime"

            If m_numKazokuIdx = 0 OrElse m_numKazokuKensu = 0 Then
                fKY_FirstTime = 0
            Else
                fKY_FirstTime = g_KazokuInfo(m_numKazokuIdx).lngFirstTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks>�Ƒ����C���f�b�N�X���O�A�܂��́A�Ƒ���񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fKY_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fKY_UpdTime"

            If m_numKazokuIdx = 0 OrElse m_numKazokuKensu = 0 Then
                fKY_UpdTime = 0
            Else
                fKY_UpdTime = g_KazokuInfo(m_numKazokuIdx).lngUpdTime
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    
    ''' <summary>
    ''' ���C�����擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetStudyInfo() As Boolean

        General.g_ErrorProc = "clsStaffIdo mGetStudyInfo"

        mGetStudyInfo = False

        Try


            '�N�x�E�폜�󋵂�������
            m_numNendo = 0
            m_intDelKbn = 2

            '�擾
            If fncGetStudyInfo() = False Then
                Exit Function
            End If

            '����I��
            mGetStudyInfo = True

        Catch
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���C���f�[�^���擾����
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function mGetStudyInfo2() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetStudyInfo2"

            mGetStudyInfo2 = False

            '�擾
            If fncGetStudyInfo() = False Then
                Exit Function
            End If

            '����I��
            mGetStudyInfo2 = True



        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���C��񌏐����擾����
    ''' </summary>
    ''' <returns>���C��񌏐�</returns>
    ''' <remarks></remarks>
    Public Function fSD_StudyCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_StudyCount"

            fSD_StudyCount = m_numStudyKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���C���t��񌏐����擾����
    ''' </summary>
    ''' <returns>���C���t��񌏐�</returns>
    ''' <remarks></remarks>
    Public Function fSD_StudyDateCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_StudyDateCount"

            m_numStudyDateKensu = UBound(g_StudyInfo(m_numStudyIdx).objDateList)
            fSD_StudyDateCount = m_numStudyDateKensu


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �N�x���擾����
    ''' </summary>
    ''' <returns>�N�x</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSD_YYYY() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_YYYY"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_YYYY = 0
            Else
                fSD_YYYY = g_StudyInfo(m_numStudyIdx).numYYYY
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ����/�V�K�ް� �����׸ނ��擾����
    ''' </summary>
    ''' <returns>����/�V�K�ް� �����׸�</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSD_NewFlg() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_NewFlg"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_NewFlg = 0
            Else
                fSD_NewFlg = g_StudyInfo(m_numStudyIdx).numNewFlg
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���� �����׸ނ��擾����
    ''' </summary>
    ''' <returns>���� �����׸�</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSD_ProcFlg() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_ProcFlg"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_ProcFlg = 0
            Else
                fSD_ProcFlg = g_StudyInfo(m_numStudyIdx).numProcFlg
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ؽĕ\�����̲��ޯ�����擾����
    ''' </summary>
    ''' <returns>ؽĕ\�����̲��ޯ��</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A0���擾����</remarks>
    Public Function fSD_DispIndex() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_DispIndex"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_DispIndex = 0
            Else
                fSD_DispIndex = g_StudyInfo(m_numStudyIdx).numDispIndex
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' SEQ No.���擾����
    ''' </summary>
    ''' <returns>SEQ No.</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_SEQ() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_SEQ"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_SEQ = ""
            Else
                fSD_SEQ = g_StudyInfo(m_numStudyIdx).strSEQ
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ������ނ��擾����
    ''' </summary>
    ''' <returns>�������</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_CourseCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_CourseCD"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_CourseCD = ""
            Else
                fSD_CourseCD = g_StudyInfo(m_numStudyIdx).strCourseCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �敪���ނ��擾����
    ''' </summary>
    ''' <returns>�敪����</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_KbnCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_KbnCD"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_KbnCD = ""
            Else
                fSD_KbnCD = g_StudyInfo(m_numStudyIdx).strKbnCD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��޺��ނ��擾����
    ''' </summary>
    ''' <returns>��޺���</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_SyuruiCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_SyuruiCD"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_SyuruiCD = ""
            Else
                fSD_SyuruiCD = g_StudyInfo(m_numStudyIdx).strSyuruiCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��ú��ނ��擾����
    ''' </summary>
    ''' <returns>��ú���</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_SyusaiCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_SyusaiCD"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_SyusaiCD = ""
            Else
                fSD_SyusaiCD = g_StudyInfo(m_numStudyIdx).strSyusaiCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Q�����ނ��擾����
    ''' </summary>
    ''' <returns>�Q������</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_SankaCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_SankaCD"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_SankaCD = ""
            Else
                fSD_SankaCD = g_StudyInfo(m_numStudyIdx).strSankaCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��u�󋵂��擾����
    ''' </summary>
    ''' <returns>��u��</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_ApplyStatus() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_ApplyStatus"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_ApplyStatus = ""
            Else
                fSD_ApplyStatus = g_StudyInfo(m_numStudyIdx).strApplyStatus
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �폜�󋵂��擾����
    ''' </summary>
    ''' <returns>�폜��</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_DeleteStatus() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_DeleteStatus"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_DeleteStatus = ""
            Else
                fSD_DeleteStatus = g_StudyInfo(m_numStudyIdx).strDeleteStatus
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��u�񍐂��擾����
    ''' </summary>
    ''' <returns>��u��</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_ApplyRepo() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_ApplyRepo"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_ApplyRepo = ""
            Else
                fSD_ApplyRepo = g_StudyInfo(m_numStudyIdx).strApplyRepo
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���l���擾����
    ''' </summary>
    ''' <returns>���l</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_Biko() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_Biko"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_Biko = ""
            Else
                fSD_Biko = g_StudyInfo(m_numStudyIdx).strBiko
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �R�[�X�@���̂��擾����
    ''' </summary>
    ''' <returns>�R�[�X�@����</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_CourseName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_CourseName"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_CourseName = ""
            Else
                fSD_CourseName = g_StudyInfo(m_numStudyIdx).strCorseName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���C�敪�@���̂��擾����
    ''' </summary>
    ''' <returns>���C�敪�@����</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_KbnName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_KbnName"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_KbnName = ""
            Else
                fSD_KbnName = g_StudyInfo(m_numStudyIdx).strKbnName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��ށ@���̂��擾����
    ''' </summary>
    ''' <returns>��ށ@����</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_SyuruiName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_SyuruiName"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_SyuruiName = ""
            Else
                fSD_SyuruiName = g_StudyInfo(m_numStudyIdx).strSyuruiName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��Á@���̂��擾����
    ''' </summary>
    ''' <returns>��Á@����</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_SyusaiName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_SyusaiName"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_SyusaiName = ""
            Else
                fSD_SyusaiName = g_StudyInfo(m_numStudyIdx).strSyusaiName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Q���@���̂��擾����
    ''' </summary>
    ''' <returns>�Q���@����</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_SankaName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_SankaName"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_SankaName = ""
            Else
                fSD_SankaName = g_StudyInfo(m_numStudyIdx).strSankaName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �e�[�}���擾����
    ''' </summary>
    ''' <returns>�e�[�}</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_Thema() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_Thema"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_Thema = ""
            Else
                fSD_Thema = g_StudyInfo(m_numStudyIdx).strThema
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �v��FLG���擾����
    ''' </summary>
    ''' <returns>�v��FLG</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_PlaningFLG() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_PlaningFLG"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_PlaningFLG = ""
            Else
                fSD_PlaningFLG = g_StudyInfo(m_numStudyIdx).strPlaningFLG
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��pCD���擾����
    ''' </summary>
    ''' <returns>��pCD</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_CostCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_CostCD"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_CostCD = ""
            Else
                fSD_CostCD = g_StudyInfo(m_numStudyIdx).strCostCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��p���̂��擾����
    ''' </summary>
    ''' <returns>��p����</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_CostName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_CostName"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_CostName = ""
            Else
                fSD_CostName = g_StudyInfo(m_numStudyIdx).strCostName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��������܂Ƃ߂ɂ������̂��擾����
    ''' </summary>
    ''' <returns>��������܂Ƃ߂ɂ�������</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_Date() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_Date"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_Date = ""
            Else
                fSD_Date = g_StudyInfo(m_numStudyIdx).strDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �J�n�����擾����
    ''' </summary>
    ''' <returns>�J�n��</returns>
    ''' <remarks>
    ''' ���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A0���擾����<br />
    ''' ��L�ȊO�ŁA���C���t���C���f�b�N�X���O�A�܂��́A���C���t��񌏐����O�̏ꍇ�A0���擾����
    ''' </remarks>
    Public Function fSD_FromDate() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_FromDate"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_FromDate = 0
            Else
                If m_numStudyDateIdx = 0 OrElse m_numStudyDateKensu = 0 Then
                    fSD_FromDate = 0
                Else
                    fSD_FromDate = g_StudyInfo(m_numStudyIdx).objDateList(m_numStudyDateIdx).numFromDate
                End If
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �I�������擾����
    ''' </summary>
    ''' <returns>�I����</returns>
    ''' <remarks>
    ''' ���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A0���擾����<br />
    ''' ��L�ȊO�ŁA���C���t���C���f�b�N�X���O�A�܂��́A���C���t��񌏐����O�̏ꍇ�A0���擾����
    ''' </remarks>
    Public Function fSD_ToDate() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_ToDate"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_ToDate = 0
            Else
                If m_numStudyDateIdx = 0 OrElse m_numStudyDateKensu = 0 Then
                    fSD_ToDate = 0
                Else
                    fSD_ToDate = g_StudyInfo(m_numStudyIdx).objDateList(m_numStudyDateIdx).numToDate
                End If
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���Ԃ̃^�C�v���擾����
    ''' </summary>
    ''' <returns>���Ԃ̃^�C�v</returns>
    ''' <remarks>
    ''' ���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����<br />
    ''' ��L�ȊO�ŁA���C���t���C���f�b�N�X���O�A�܂��́A���C���t��񌏐����O�̏ꍇ�A""���擾����
    ''' </remarks>
    Public Function fSD_DateType() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_DateType"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_DateType = ""
            Else
                If m_numStudyDateIdx = 0 OrElse m_numStudyDateKensu = 0 Then
                    fSD_DateType = ""
                Else
                    fSD_DateType = g_StudyInfo(m_numStudyIdx).objDateList(m_numStudyDateIdx).strDateType
                End If
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��pCD2(���C�\��F)���擾����
    ''' </summary>
    ''' <returns>��pCD2(���C�\��F)</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_CostCD2() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_CostCD2"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_CostCD2 = ""
            Else
                fSD_CostCD2 = g_StudyInfo(m_numStudyIdx).strCostCD2
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��p����2(���C�\��F)���擾����
    ''' </summary>
    ''' <returns>��p����2(���C�\��F)</returns>
    ''' <remarks>���C���C���f�b�N�X���O�A�܂��́A���C��񌏐����O�̏ꍇ�A""���擾����</remarks>
    Public Function fSD_CostName2() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fSD_CostName2"

            If m_numStudyIdx = 0 OrElse m_numStudyKensu = 0 Then
                fSD_CostName2 = ""
            Else
                fSD_CostName2 = g_StudyInfo(m_numStudyIdx).strCostName2
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Ɛя��f�a�@�R�[�h���擾����
    ''' </summary>
    ''' <returns>�Ɛя��f�a�@�R�[�h</returns>
    ''' <remarks>�ƐуC���f�b�N�X���O�A�܂��́A�Ɛь������O�̏ꍇ�A""���擾����</remarks>
    Public Function fGY_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGY_HospitalCD"

            If m_numGyosekiIdx = 0 OrElse m_numGyosekiKensu = 0 Then
                fGY_HospitalCD = ""
            Else
                fGY_HospitalCD = g_Gyoseki(m_numGyosekiIdx).strHospitalCD
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �Ɛя��f�E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�Ɛя��f�E���Ǘ��ԍ�</returns>
    ''' <remarks>�ƐуC���f�b�N�X���O�A�܂��́A�Ɛь������O�̏ꍇ�A""���擾����</remarks>
    Public Function fGY_StaffMngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGY_StaffMngID"

            If m_numGyosekiIdx = 0 OrElse m_numGyosekiKensu = 0 Then
                fGY_StaffMngID = ""
            Else
                fGY_StaffMngID = g_Gyoseki(m_numGyosekiIdx).strStaffMngID
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �Ɛя��f�������擾����
    ''' </summary>
    ''' <returns>�Ɛя��f����</returns>
    ''' <remarks></remarks>
    Public Function fGY_GyosekiCount() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fGY_GyosekiCount"

            fGY_GyosekiCount = m_numGyosekiKensu


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �Ɛя��f�ƐуR�[�h���擾����
    ''' </summary>
    ''' <returns>�Ɛя��f�ƐуR�[�h</returns>
    ''' <remarks>�ƐуC���f�b�N�X���O�A�܂��́A�Ɛь������O�̏ꍇ�A""���擾����</remarks>
    Public Function fGY_GyosekiCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGY_GyosekiCD"

            If m_numGyosekiIdx = 0 OrElse m_numGyosekiKensu = 0 Then
                fGY_GyosekiCD = ""
            Else
                fGY_GyosekiCD = g_Gyoseki(m_numGyosekiIdx).strGyosekiCd
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �Ɛя��f�Ɛі��̂��擾����
    ''' </summary>
    ''' <returns>�Ɛя��f�Ɛі���</returns>
    ''' <remarks>�ƐуC���f�b�N�X���O�A�܂��́A�Ɛь������O�̏ꍇ�A""���擾����</remarks>
    Public Function fGY_GyosekiName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGY_GyosekiName"

            If m_numGyosekiIdx = 0 OrElse m_numGyosekiKensu = 0 Then
                fGY_GyosekiName = ""
            Else
                fGY_GyosekiName = g_Gyoseki(m_numGyosekiIdx).strGyosekiName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �Ɛя��f�J�n�N�������擾����
    ''' </summary>
    ''' <returns>�Ɛя��f�J�n�N����</returns>
    ''' <remarks>�ƐуC���f�b�N�X���O�A�܂��́A�Ɛь������O�̏ꍇ�A0���擾����</remarks>
    Public Function fGY_FromDate() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fGY_FromDate"

            If m_numGyosekiIdx = 0 OrElse m_numGyosekiKensu = 0 Then
                fGY_FromDate = 0
            Else
                fGY_FromDate = g_Gyoseki(m_numGyosekiIdx).numFromDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �Ɛя��f�I���N�������擾����
    ''' </summary>
    ''' <returns>�Ɛя��f�I���N����</returns>
    ''' <remarks>�ƐуC���f�b�N�X���O�A�܂��́A�Ɛь������O�̏ꍇ�A0���擾����</remarks>
    Public Function fGY_ToDate() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fGY_ToDate"

            If m_numGyosekiIdx = 0 OrElse m_numGyosekiKensu = 0 Then
                fGY_ToDate = 0
            Else
                fGY_ToDate = Integer.Parse(g_Gyoseki(m_numGyosekiIdx).numToDate)
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �Ɛя��f������擾����
    ''' </summary>
    ''' <returns>�Ɛя��f����</returns>
    ''' <remarks>�ƐуC���f�b�N�X���O�A�܂��́A�Ɛь������O�̏ꍇ�A""���擾����</remarks>
    Public Function fGY_Subject() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGY_Subject"

            If m_numGyosekiIdx = 0 OrElse m_numGyosekiKensu = 0 Then
                fGY_Subject = ""
            Else
                fGY_Subject = g_Gyoseki(m_numGyosekiIdx).strSubject
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �Ɛя��f�Ɛє��\�ꏊ�R�[�h���擾����
    ''' </summary>
    ''' <returns>�Ɛя��f�Ɛє��\�ꏊ�R�[�h</returns>
    ''' <remarks>�ƐуC���f�b�N�X���O�A�܂��́A�Ɛь������O�̏ꍇ�A""���擾����</remarks>
    Public Function fGY_GyosekiPlaceCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGY_GyosekiPlaceCD"

            If m_numGyosekiIdx = 0 OrElse m_numGyosekiKensu = 0 Then
                fGY_GyosekiPlaceCD = ""
            Else
                fGY_GyosekiPlaceCD = g_Gyoseki(m_numGyosekiIdx).strGyosekiPlaceCd
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �Ɛя��f�Ɛє��\�ꏊ���̂��擾����
    ''' </summary>
    ''' <returns>�Ɛя��f�Ɛє��\�ꏊ����</returns>
    ''' <remarks>�ƐуC���f�b�N�X���O�A�܂��́A�Ɛь������O�̏ꍇ�A""���擾����</remarks>
    Public Function fGY_GyosekiPlaceName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGY_GyosekiPlaceName"

            If m_numGyosekiIdx = 0 OrElse m_numGyosekiKensu = 0 Then
                fGY_GyosekiPlaceName = ""
            Else
                fGY_GyosekiPlaceName = g_Gyoseki(m_numGyosekiIdx).strGyosekiPlaceName
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �Ɛя��f�Ɛє��l���擾����
    ''' </summary>
    ''' <returns>�Ɛя��f�Ɛє��l</returns>
    ''' <remarks>�ƐуC���f�b�N�X���O�A�܂��́A�Ɛь������O�̏ꍇ�A""���擾����</remarks>
    Public Function fGY_GyosekiBikou() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fGY_GyosekiBikou"

            If m_numGyosekiIdx = 0 OrElse m_numGyosekiKensu = 0 Then
                fGY_GyosekiBikou = ""
            Else
                fGY_GyosekiBikou = g_Gyoseki(m_numGyosekiIdx).strGyosekiBikou
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �Ɛя��f����o�^�������擾����
    ''' </summary>
    ''' <returns>�Ɛя��f����o�^����</returns>
    ''' <remarks>�ƐуC���f�b�N�X���O�A�܂��́A�Ɛь������O�̏ꍇ�A0���擾����</remarks>
    Public Function fGY_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fGY_FirstTime"

            If m_numGyosekiIdx = 0 OrElse m_numGyosekiKensu = 0 Then
                fGY_FirstTime = 0
            Else
                fGY_FirstTime = g_Gyoseki(m_numGyosekiIdx).dblRegistFirstTimeDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �Ɛя��f�ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�Ɛя��f�ŏI�X�V����</returns>
    ''' <remarks>�ƐуC���f�b�N�X���O�A�܂��́A�Ɛь������O�̏ꍇ�A0���擾����</remarks>
    Public Function fGY_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fGY_UpdTime"

            If m_numGyosekiIdx = 0 OrElse m_numGyosekiKensu = 0 Then
                fGY_UpdTime = 0
            Else
                fGY_UpdTime = g_Gyoseki(m_numGyosekiIdx).dblLastUpdTimeDate
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' SEQ���擾����
    ''' </summary>
    ''' <returns>SEQ</returns>
    ''' <remarks>�ƐуC���f�b�N�X���O�A�܂��́A�Ɛь������O�̏ꍇ�A0���擾����</remarks>
    Public Function fGY_Seq() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fGY_Seq"

            If m_numGyosekiIdx = 0 OrElse m_numGyosekiKensu = 0 Then
                fGY_Seq = 0
            Else
                fGY_Seq = g_Gyoseki(m_numGyosekiIdx).numSEQ
            End If


        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' �̗p�ٓ������擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetSaiyoIdo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetSaiyoIdo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Dim w_Rs As ADODB.Recordset
        Dim w_EmpCD_F As ADODB.Field
        Dim w_EmpName_F As ADODB.Field
        Dim w_EmpSecName_F As ADODB.Field
        Dim w_EmpDate_F As ADODB.Field
        Dim w_RetireCD_F As ADODB.Field
        Dim w_RetireName_F As ADODB.Field
        Dim w_RetireDate_F As ADODB.Field
        Dim w_FirstTime_F As ADODB.Field
        Dim w_UpdTime_F As ADODB.Field
        Dim w_StaffID_F As ADODB.Field

        Const w_EMP As String = "C001"
        Const w_RETIRE As String = "C002"

        fncGetSaiyoIdo = False
        Try
            ReDim g_SaiyoIdo.Ido(0)

            w_strSql = ""
            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, CStr(General.gInstall_Enum.AccessType_PassThrough)).Equals(CStr(General.gInstall_Enum.AccessType_PassThrough)) Then 'ORACLE

                w_strSql = w_strSql & " SELECT SM.EMPCD "
                w_strSql = w_strSql & " ,      SM.EMPDATE "
                w_strSql = w_strSql & " ,      H1.NAME      AS EMPNAME "
                w_strSql = w_strSql & " ,      H1.SECNAME   AS EMPSECNAME "
                w_strSql = w_strSql & " ,      SM.RETIRECD "
                w_strSql = w_strSql & " ,      SM.RETIREDATE "
                w_strSql = w_strSql & " ,      H2.NAME      AS RETIRENAME "
                w_strSql = w_strSql & " ,      SM.REGISTFIRSTTIMEDATE "
                w_strSql = w_strSql & " ,      SM.LASTUPDTIMEDATE "
                w_strSql = w_strSql & " ,      SM.STAFFID "
                w_strSql = w_strSql & " FROM   NS_STAFFMNGHISTORY_F SM "
                w_strSql = w_strSql & " ,      NS_HANYOU_M          H1 "
                w_strSql = w_strSql & " ,      NS_HANYOU_M          H2 "
                w_strSql = w_strSql & " WHERE  SM.HOSPITALCD     = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    SM.STAFFMNGID     = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    H1.HOSPITALCD (+) = SM.HOSPITALCD "
                w_strSql = w_strSql & " AND    H1.MASTERID (+)   = '" & w_EMP & "' "
                w_strSql = w_strSql & " AND    H1.MASTERCD (+)   = SM.EMPCD "
                w_strSql = w_strSql & " AND    H2.HOSPITALCD (+) = SM.HOSPITALCD "
                w_strSql = w_strSql & " AND    H2.MASTERID (+)   = '" & w_RETIRE & "' "
                w_strSql = w_strSql & " AND    H2.MASTERCD (+)   = SM.RETIRECD "

            Else '����ȊO

                w_strSql = w_strSql & " SELECT SM.EMPCD "
                w_strSql = w_strSql & " ,      SM.EMPDATE "
                w_strSql = w_strSql & " ,      H1.NAME      AS EMPNAME "
                w_strSql = w_strSql & " ,      H1.SECNAME   AS EMPSECNAME "
                w_strSql = w_strSql & " ,      SM.RETIRECD "
                w_strSql = w_strSql & " ,      SM.RETIREDATE "
                w_strSql = w_strSql & " ,      H2.NAME      AS RETIRENAME "
                w_strSql = w_strSql & " ,      SM.REGISTFIRSTTIMEDATE "
                w_strSql = w_strSql & " ,      SM.LASTUPDTIMEDATE "
                w_strSql = w_strSql & " ,      SM.STAFFID "
                w_strSql = w_strSql & " FROM   NS_STAFFMNGHISTORY_F SM "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M H1 "
                w_strSql = w_strSql & " ON   H1.HOSPITALCD = SM.HOSPITALCD "
                w_strSql = w_strSql & " AND  H1.MASTERID   = '" & w_EMP & "' "
                w_strSql = w_strSql & " AND  H1.MASTERCD   = SM.EMPCD "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M H2 "
                w_strSql = w_strSql & " ON   H2.HOSPITALCD = SM.HOSPITALCD "
                w_strSql = w_strSql & " AND  H2.MASTERID   = '" & w_RETIRE & "' "
                w_strSql = w_strSql & " AND  H2.MASTERCD   = SM.RETIRECD "
                w_strSql = w_strSql & " WHERE  SM.HOSPITALCD     = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    SM.STAFFMNGID     = '" & m_strStaffMngID & "' "

            End If
            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND SM.EMPDATE      <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND   SM.EMPDATE    <= " & m_numDateTo & " "
                w_strSql = w_strSql & " AND ( SM.RETIREDATE >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " OR    SM.RETIREDATE  = 0 "
                w_strSql = w_strSql & " OR    SM.RETIREDATE IS NULL ) "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY SM.EMPDATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY SM.EMPDATE DESC "
            End If


            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numSaiyoKensu = 0
                    .Close()
                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numSaiyoKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_SaiyoIdo.Ido(m_numSaiyoKensu)

                    w_EmpCD_F = .Fields("EMPCD")
                    w_EmpName_F = .Fields("EMPNAME")
                    w_EmpSecName_F = .Fields("EMPSECNAME")
                    w_EmpDate_F = .Fields("EMPDATE")
                    w_RetireCD_F = .Fields("RETIRECD")
                    w_RetireName_F = .Fields("RETIRENAME")
                    w_RetireDate_F = .Fields("RETIREDATE")
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE")
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE")
                    w_StaffID_F = .Fields("STAFFID")

                    g_SaiyoIdo.strHospitalCD = m_strHospitalCD
                    g_SaiyoIdo.strStaffMngID = m_strStaffMngID

                    For w_numLoop = 1 To m_numSaiyoKensu

                        '�̗p�R�[�h
                        g_SaiyoIdo.Ido(w_numLoop).strEmpCD = General.paGetDbFieldVal(w_EmpCD_F, "")
                        '�̗p����
                        g_SaiyoIdo.Ido(w_numLoop).strEmpName = General.paGetDbFieldVal(w_EmpName_F, "")
                        '�̗p����
                        g_SaiyoIdo.Ido(w_numLoop).strEmpSecName = General.paGetDbFieldVal(w_EmpSecName_F, "")
                        '�̗p��
                        g_SaiyoIdo.Ido(w_numLoop).numEmpDate = Integer.Parse(General.paGetDbFieldVal(w_EmpDate_F, 0))
                        '�ސE�R�[�h
                        g_SaiyoIdo.Ido(w_numLoop).strRetireCD = General.paGetDbFieldVal(w_RetireCD_F, "")
                        '�ސE����
                        g_SaiyoIdo.Ido(w_numLoop).strRetireName = General.paGetDbFieldVal(w_RetireName_F, "")
                        '�ސE��
                        g_SaiyoIdo.Ido(w_numLoop).numRetireDate = Integer.Parse(General.paGetDbFieldVal(w_RetireDate_F, 99999999))
                        '����o�^����
                        g_SaiyoIdo.Ido(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_SaiyoIdo.Ido(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))
                        '�E���ԍ�
                        g_SaiyoIdo.Ido(w_numLoop).strStaffID = General.paGetDbFieldVal(w_StaffID_F, "")

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With
            w_Rs = Nothing

            fncGetSaiyoIdo = True

            General.g_ErrorProc = w_strPreErrorProc
        Catch
            Throw
        End Try
    End Function


    ''' <summary>
    ''' �Ζ������ٓ������擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetKinmuDeptIdo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetKinmuDeptIdo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Dim w_Rs As ADODB.Recordset
        Dim w_CD_F As ADODB.Field
        Dim w_Name_F As ADODB.Field
        Dim w_DateFrom_F As ADODB.Field
        Dim w_DateTo_F As ADODB.Field
        Dim w_FirstTime_F As ADODB.Field
        Dim w_UpdTime_F As ADODB.Field
        Dim w_IdoHope_F As ADODB.Field
        Dim w_SecName_F As ADODB.Field
        Dim w_DispNo_F As ADODB.Field
        Const w_IDOHOPE As String = "��]"
        Dim w_strUseNaviFlg As String 'NAVI�^�pFLG

        fncGetKinmuDeptIdo = False
        Try
            ReDim g_KinmuDeptIdo.Ido(0)

            'NAVI�^�pFLG�擾
            w_strUseNaviFlg = General.paGetItemValue(General.G_StrMainKey1, General.G_StrSection1, "USENAVIFLG", "0", m_strHospitalCD)

            w_strSql = ""
            w_strSql = w_strSql & " SELECT KI.KINMUDEPTCD "
            w_strSql = w_strSql & " ,      KD.NAME "
            w_strSql = w_strSql & " ,      KI.IDODATE "
            w_strSql = w_strSql & " ,      KI.ENDDATE "
            w_strSql = w_strSql & " ,      KI.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      KI.LASTUPDTIMEDATE "

            'NAVI�^�p�̏ꍇ�A�ٓ���]�t���O���擾
            If w_strUseNaviFlg = "1" Then
                w_strSql = w_strSql & " ,      KI.IDOHOPEFLG "
            End If

            w_strSql = w_strSql & " ,      KD.SECNAME "
            w_strSql = w_strSql & " ,      KD.DISPNO "

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, CStr(General.gInstall_Enum.AccessType_PassThrough)).Equals(CStr(General.gInstall_Enum.AccessType_PassThrough)) Then 'ORACLE

                w_strSql = w_strSql & " FROM   NS_KINMUIDOINFO_F KI "
                w_strSql = w_strSql & " ,      NS_KINMUDEPT_M    KD "
                w_strSql = w_strSql & " WHERE  KI.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    KI.STAFFMNGID    = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    KD.HOSPITALCD  (+)  = KI.HOSPITALCD "
                w_strSql = w_strSql & " AND    KD.KINMUDEPTCD (+)  = KI.KINMUDEPTCD "

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_KINMUIDOINFO_F KI "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_KINMUDEPT_M KD "
                w_strSql = w_strSql & " ON   KD.HOSPITALCD  = KI.HOSPITALCD "
                w_strSql = w_strSql & " AND  KD.KINMUDEPTCD = KI.KINMUDEPTCD "
                w_strSql = w_strSql & " WHERE  KI.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    KI.STAFFMNGID    = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND KI.IDODATE      <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND   KI.IDODATE    <= " & m_numDateTo & " "
                w_strSql = w_strSql & " AND ( KI.ENDDATE >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " OR    KI.ENDDATE  = 0 "
                w_strSql = w_strSql & " OR    KI.ENDDATE IS NULL ) "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY KI.IDODATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY KI.IDODATE DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numKinmuDeptKensu = 0
                    .Close()
                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numKinmuDeptKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_KinmuDeptIdo.Ido(m_numKinmuDeptKensu)

                    w_CD_F = .Fields("KINMUDEPTCD")
                    w_Name_F = .Fields("NAME")
                    w_DateFrom_F = .Fields("IDODATE")
                    w_DateTo_F = .Fields("ENDDATE")
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE")
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE")
                    If w_strUseNaviFlg = "1" Then 'NAVI�^�p�̏ꍇ�͒ǉ�
                        w_IdoHope_F = .Fields("IDOHOPEFLG")
                    End If
                    w_SecName_F = .Fields("SECNAME")
                    w_DispNo_F = .Fields("DISPNO")

                    g_KinmuDeptIdo.strHospitalCD = m_strHospitalCD
                    g_KinmuDeptIdo.strStaffMngID = m_strStaffMngID

                    For w_numLoop = 1 To m_numKinmuDeptKensu
                        '�R�[�h
                        g_KinmuDeptIdo.Ido(w_numLoop).strCD = General.paGetDbFieldVal(w_CD_F, "")
                        '����
                        g_KinmuDeptIdo.Ido(w_numLoop).strName = General.paGetDbFieldVal(w_Name_F, "")
                        '�J�n��
                        g_KinmuDeptIdo.Ido(w_numLoop).numDateFrom = Integer.Parse(General.paGetDbFieldVal(w_DateFrom_F, 0))
                        '�I����
                        g_KinmuDeptIdo.Ido(w_numLoop).numDateTo = Integer.Parse(General.paGetDbFieldVal(w_DateTo_F, 99999999))
                        '����o�^����
                        g_KinmuDeptIdo.Ido(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_KinmuDeptIdo.Ido(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))

                        If "1".Equals(w_strUseNaviFlg) Then 'NAVI�^�p�̏ꍇ�͒ǉ�
                            '�ٓ���]�t���O
                            If "1".Equals(w_IdoHope_F.Value) Then
                                g_KinmuDeptIdo.Ido(w_numLoop).strIdoHope = w_IDOHOPE
                            Else
                                g_KinmuDeptIdo.Ido(w_numLoop).strIdoHope = ""
                            End If
                        End If

                        '����
                        g_KinmuDeptIdo.Ido(w_numLoop).SecName = General.paGetDbFieldVal(w_SecName_F, "")
                        '�\����
                        g_KinmuDeptIdo.Ido(w_numLoop).DispNo = General.paGetDbFieldVal(w_DispNo_F, 0)

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With
            w_Rs = Nothing

            fncGetKinmuDeptIdo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' �Čf�����ٓ������擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetSaikeiIdo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetSaikeiIdo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Dim w_Rs As ADODB.Recordset
        Dim w_CD_F As ADODB.Field
        Dim w_Name_F As ADODB.Field
        Dim w_DateFrom_F As ADODB.Field
        Dim w_DateTo_F As ADODB.Field
        Dim w_FirstTime_F As ADODB.Field
        Dim w_UpdTime_F As ADODB.Field


        fncGetSaikeiIdo = False
        Try
            ReDim g_SaikeiIdo.Ido(0)

            w_strSql = ""
            w_strSql = w_strSql & " SELECT SI.SAIKEICD "
            w_strSql = w_strSql & " ,      KD.NAME "
            w_strSql = w_strSql & " ,      SI.IDODATE "
            w_strSql = w_strSql & " ,      SI.ENDDATE "
            w_strSql = w_strSql & " ,      SI.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      SI.LASTUPDTIMEDATE "
            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, CStr(General.gInstall_Enum.AccessType_PassThrough)).Equals(CStr(General.gInstall_Enum.AccessType_PassThrough)) Then 'ORACLE

                w_strSql = w_strSql & " FROM   NS_SAIKEIIDOINFO_F SI "
                w_strSql = w_strSql & " ,      NS_KINMUDEPT_M    KD "
                w_strSql = w_strSql & " WHERE  SI.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    SI.STAFFMNGID    = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    KD.HOSPITALCD  (+)  = SI.HOSPITALCD "
                w_strSql = w_strSql & " AND    KD.KINMUDEPTCD (+)  = SI.SAIKEICD "

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_SAIKEIIDOINFO_F SI "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_KINMUDEPT_M KD "
                w_strSql = w_strSql & " ON   KD.HOSPITALCD  = SI.HOSPITALCD "
                w_strSql = w_strSql & " AND  KD.KINMUDEPTCD = SI.SAIKEICD "
                w_strSql = w_strSql & " WHERE  SI.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    SI.STAFFMNGID    = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND SI.IDODATE      <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND   SI.IDODATE    <= " & m_numDateTo & " "
                w_strSql = w_strSql & " AND ( SI.ENDDATE >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " OR    SI.ENDDATE  = 0 "
                w_strSql = w_strSql & " OR    SI.ENDDATE IS NULL ) "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY SI.IDODATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY SI.IDODATE DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numSaikeiKensu = 0
                    .Close()
                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numSaikeiKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_SaikeiIdo.Ido(m_numSaikeiKensu)

                    w_CD_F = .Fields("SAIKEICD")
                    w_Name_F = .Fields("NAME")
                    w_DateFrom_F = .Fields("IDODATE")
                    w_DateTo_F = .Fields("ENDDATE")
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE")
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE")

                    g_SaikeiIdo.strHospitalCD = m_strHospitalCD
                    g_SaikeiIdo.strStaffMngID = m_strStaffMngID

                    For w_numLoop = 1 To m_numSaikeiKensu
                        '�R�[�h
                        g_SaikeiIdo.Ido(w_numLoop).strCD = General.paGetDbFieldVal(w_CD_F, "")
                        '����
                        g_SaikeiIdo.Ido(w_numLoop).strName = General.paGetDbFieldVal(w_Name_F, "")
                        '�J�n��
                        g_SaikeiIdo.Ido(w_numLoop).numDateFrom = Integer.Parse(General.paGetDbFieldVal(w_DateFrom_F, 0))
                        '�I����
                        g_SaikeiIdo.Ido(w_numLoop).numDateTo = Integer.Parse(General.paGetDbFieldVal(w_DateTo_F, 99999999))
                        '����o�^����
                        g_SaikeiIdo.Ido(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_SaikeiIdo.Ido(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetSaikeiIdo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �z�������ٓ������擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetWardDeptIdo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetWardDeptIdo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Dim w_Rs As ADODB.Recordset
        Dim w_CD_F As ADODB.Field
        Dim w_Name_F As ADODB.Field
        Dim w_DateFrom_F As ADODB.Field
        Dim w_DateTo_F As ADODB.Field
        Dim w_FirstTime_F As ADODB.Field
        Dim w_UpdTime_F As ADODB.Field


        fncGetWardDeptIdo = False
        Try
            ReDim g_WardDeptIdo.Ido(0)

            w_strSql = ""
            w_strSql = w_strSql & " SELECT WI.WARDDEPTCD "
            w_strSql = w_strSql & " ,      WD.NAME "
            w_strSql = w_strSql & " ,      WI.IDODATE "
            w_strSql = w_strSql & " ,      WI.ENDDATE "
            w_strSql = w_strSql & " ,      WI.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      WI.LASTUPDTIMEDATE "

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals(CStr(General.gInstall_Enum.AccessType_PassThrough)) Then 'ORACLE

                w_strSql = w_strSql & " FROM   NS_WARDIDOINFO_F WI "
                w_strSql = w_strSql & " ,      NS_WARDDEPT_M    WD "
                w_strSql = w_strSql & " WHERE  WI.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    WI.STAFFMNGID    = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    WD.HOSPITALCD (+)  = WI.HOSPITALCD "
                w_strSql = w_strSql & " AND    WD.WARDDEPTCD (+)  = WI.WARDDEPTCD "

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_WARDIDOINFO_F WI "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_WARDDEPT_M WD "
                w_strSql = w_strSql & " ON     WD.HOSPITALCD    = WI.HOSPITALCD "
                w_strSql = w_strSql & " AND    WD.WARDDEPTCD    = WI.WARDDEPTCD "
                w_strSql = w_strSql & " WHERE  WI.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    WI.STAFFMNGID    = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND WI.IDODATE      <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND   WI.IDODATE    <= " & m_numDateTo & " "
                w_strSql = w_strSql & " AND ( WI.ENDDATE >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " OR    WI.ENDDATE  = 0 "
                w_strSql = w_strSql & " OR    WI.ENDDATE IS NULL ) "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY WI.IDODATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY WI.IDODATE DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numWardDeptKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numWardDeptKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_WardDeptIdo.Ido(m_numWardDeptKensu)

                    w_CD_F = .Fields("WARDDEPTCD")
                    w_Name_F = .Fields("NAME")
                    w_DateFrom_F = .Fields("IDODATE")
                    w_DateTo_F = .Fields("ENDDATE")
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE")
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE")

                    g_WardDeptIdo.strHospitalCD = m_strHospitalCD
                    g_WardDeptIdo.strStaffMngID = m_strStaffMngID

                    For w_numLoop = 1 To m_numWardDeptKensu
                        '�R�[�h
                        g_WardDeptIdo.Ido(w_numLoop).strCD = General.paGetDbFieldVal(w_CD_F, "")
                        '����
                        g_WardDeptIdo.Ido(w_numLoop).strName = General.paGetDbFieldVal(w_Name_F, "")
                        '�J�n��
                        g_WardDeptIdo.Ido(w_numLoop).numDateFrom = Integer.Parse(General.paGetDbFieldVal(w_DateFrom_F, 0))
                        '�I����
                        g_WardDeptIdo.Ido(w_numLoop).numDateTo = Integer.Parse(General.paGetDbFieldVal(w_DateTo_F, 99999999))
                        '����o�^����
                        g_WardDeptIdo.Ido(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_WardDeptIdo.Ido(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetWardDeptIdo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��E�ٓ������擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetPostIdo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetPostIdo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Dim w_Rs As ADODB.Recordset
        Dim w_CD_F As ADODB.Field
        Dim w_Name_F As ADODB.Field
        Dim w_SecName_F As ADODB.Field
        Dim w_DateFrom_F As ADODB.Field
        Dim w_DateTo_F As ADODB.Field
        Dim w_FirstTime_F As ADODB.Field
        Dim w_UpdTime_F As ADODB.Field


        fncGetPostIdo = False
        Try
            ReDim g_PostIdo.Ido(0)

            w_strSql = ""
            w_strSql = w_strSql & " SELECT PI.POSTCD "
            w_strSql = w_strSql & " ,      PD.NAME "
            w_strSql = w_strSql & " ,      PD.SECNAME "
            w_strSql = w_strSql & " ,      PI.IDODATE "
            w_strSql = w_strSql & " ,      PI.ENDDATE "
            w_strSql = w_strSql & " ,      PI.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      PI.LASTUPDTIMEDATE "

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, CStr(General.gInstall_Enum.AccessType_PassThrough)).Equals(CStr(General.gInstall_Enum.AccessType_PassThrough)) Then 'ORACLE

                w_strSql = w_strSql & " FROM   NS_POSTIDOINFO_F PI "
                w_strSql = w_strSql & " ,      NS_POST_M    PD "
                w_strSql = w_strSql & " WHERE  PI.HOSPITALCD     = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    PI.STAFFMNGID     = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    PD.HOSPITALCD (+) = PI.HOSPITALCD "
                w_strSql = w_strSql & " AND    PD.POSTCD     (+) = PI.POSTCD "

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_POSTIDOINFO_F PI "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_POST_M PD "
                w_strSql = w_strSql & " ON     PD.HOSPITALCD = PI.HOSPITALCD "
                w_strSql = w_strSql & " AND    PD.POSTCD     = PI.POSTCD "
                w_strSql = w_strSql & " WHERE  PI.HOSPITALCD = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    PI.STAFFMNGID = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND PI.IDODATE      <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND   PI.IDODATE    <= " & m_numDateTo & " "
                w_strSql = w_strSql & " AND ( PI.ENDDATE >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " OR    PI.ENDDATE  = 0 "
                w_strSql = w_strSql & " OR    PI.ENDDATE IS NULL ) "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY PI.IDODATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY PI.IDODATE DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numPostKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numPostKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_PostIdo.Ido(m_numPostKensu)

                    w_CD_F = .Fields("POSTCD")
                    w_Name_F = .Fields("NAME")
                    w_SecName_F = .Fields("SECNAME")
                    w_DateFrom_F = .Fields("IDODATE")
                    w_DateTo_F = .Fields("ENDDATE")
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE")
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE")

                    g_PostIdo.strHospitalCD = m_strHospitalCD
                    g_PostIdo.strStaffMngID = m_strStaffMngID

                    For w_numLoop = 1 To m_numPostKensu
                        '�R�[�h
                        g_PostIdo.Ido(w_numLoop).strCD = General.paGetDbFieldVal(w_CD_F, "")
                        '����
                        g_PostIdo.Ido(w_numLoop).strName = General.paGetDbFieldVal(w_Name_F, "")
                        '����
                        g_PostIdo.Ido(w_numLoop).SecName = General.paGetDbFieldVal(w_SecName_F, "")
                        '�J�n��
                        g_PostIdo.Ido(w_numLoop).numDateFrom = Integer.Parse(General.paGetDbFieldVal(w_DateFrom_F, 0))
                        '�I����
                        g_PostIdo.Ido(w_numLoop).numDateTo = Integer.Parse(General.paGetDbFieldVal(w_DateTo_F, 99999999))
                        '����o�^����
                        g_PostIdo.Ido(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_PostIdo.Ido(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetPostIdo = True

            General.g_ErrorProc = w_strPreErrorProc
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �E��ٓ������擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetJobIdo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetJobIdo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Dim w_Rs As ADODB.Recordset
        Dim w_CD_F As ADODB.Field
        Dim w_Name_F As ADODB.Field
        Dim w_SecName_F As ADODB.Field
        Dim w_DateFrom_F As ADODB.Field
        Dim w_DateTo_F As ADODB.Field
        Dim w_FirstTime_F As ADODB.Field
        Dim w_UpdTime_F As ADODB.Field


        fncGetJobIdo = False
        Try
            ReDim g_JobIdo.Ido(0)

            w_strSql = ""
            w_strSql = w_strSql & " SELECT JI.JOBCD "
            w_strSql = w_strSql & " ,      JD.NAME "
            w_strSql = w_strSql & " ,      JD.SECNAME "
            w_strSql = w_strSql & " ,      JI.IDODATE "
            w_strSql = w_strSql & " ,      JI.ENDDATE "
            w_strSql = w_strSql & " ,      JI.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      JI.LASTUPDTIMEDATE "

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals((General.gInstall_Enum.AccessType_PassThrough).ToString) Then 'ORACLE

                w_strSql = w_strSql & " FROM   NS_JOBIDOINFO_F JI "
                w_strSql = w_strSql & " ,      NS_JOB_M    JD "
                w_strSql = w_strSql & " WHERE  JI.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    JI.STAFFMNGID    = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    JD.HOSPITALCD  (+)  = JI.HOSPITALCD "
                w_strSql = w_strSql & " AND    JD.JOBCD (+)  = JI.JOBCD "

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_JOBIDOINFO_F JI "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_JOB_M JD "
                w_strSql = w_strSql & " ON     JD.HOSPITALCD = JI.HOSPITALCD "
                w_strSql = w_strSql & " AND    JD.JOBCD      = JI.JOBCD "
                w_strSql = w_strSql & " WHERE  JI.HOSPITALCD = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    JI.STAFFMNGID = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND JI.IDODATE      <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND   JI.IDODATE    <= " & m_numDateTo & " "
                w_strSql = w_strSql & " AND ( JI.ENDDATE >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " OR    JI.ENDDATE  = 0 "
                w_strSql = w_strSql & " OR    JI.ENDDATE IS NULL ) "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY JI.IDODATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY JI.IDODATE DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numJobKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numJobKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_JobIdo.Ido(m_numJobKensu)

                    w_CD_F = .Fields("JOBCD")
                    w_Name_F = .Fields("NAME")
                    w_SecName_F = .Fields("SECNAME")
                    w_DateFrom_F = .Fields("IDODATE")
                    w_DateTo_F = .Fields("ENDDATE")
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE")
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE")

                    g_JobIdo.strHospitalCD = m_strHospitalCD
                    g_JobIdo.strStaffMngID = m_strStaffMngID

                    For w_numLoop = 1 To m_numJobKensu
                        '�R�[�h
                        g_JobIdo.Ido(w_numLoop).strCD = General.paGetDbFieldVal(w_CD_F, "")
                        '����
                        g_JobIdo.Ido(w_numLoop).strName = General.paGetDbFieldVal(w_Name_F, "")
                        '����
                        g_JobIdo.Ido(w_numLoop).SecName = General.paGetDbFieldVal(w_SecName_F, "")
                        '�J�n��
                        g_JobIdo.Ido(w_numLoop).numDateFrom = Integer.Parse(General.paGetDbFieldVal(w_DateFrom_F, 0))
                        '�I����
                        g_JobIdo.Ido(w_numLoop).numDateTo = Integer.Parse(General.paGetDbFieldVal(w_DateTo_F, 99999999))
                        '����o�^����
                        g_JobIdo.Ido(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_JobIdo.Ido(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetJobIdo = True

            General.g_ErrorProc = w_strPreErrorProc
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �����ٓ������擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetKenmuIdo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetKenmuIdo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Dim w_Rs As ADODB.Recordset
        Dim w_WardDeptCD_F As ADODB.Field '�z�������R�[�h
        Dim w_WardDeptName_F As ADODB.Field '�z����������
        Dim w_KinmuDeptCD_F As ADODB.Field '�Ζ������R�[�h
        Dim w_KinmuDeptName_F As ADODB.Field '�Ζ���������
        Dim w_PostCD_F As ADODB.Field '��E�R�[�h
        Dim w_PostName_F As ADODB.Field '��E����
        Dim w_DateFrom_F As ADODB.Field '�J�n�N����
        Dim w_SEQ_F As ADODB.Field 'SEQ
        Dim w_DateTo_F As ADODB.Field '�I���N����
        Dim w_FirstTime_F As ADODB.Field '����o�^����
        Dim w_UpdTime_F As ADODB.Field '�ŏI�X�V����


        fncGetKenmuIdo = False
        Try
            ReDim g_KenmuIdo.Ido(0)

            w_strSql = ""
            w_strSql = w_strSql & " SELECT KE.WARDDEPTCD "
            w_strSql = w_strSql & " ,      WM.NAME AS WARDDEPTNAME "
            w_strSql = w_strSql & " ,      KE.KINMUDEPTCD "
            w_strSql = w_strSql & " ,      KM.NAME AS KINMUDEPTNAME "
            w_strSql = w_strSql & " ,      KE.POSTCD "
            w_strSql = w_strSql & " ,      PM.NAME AS POSTNAME "
            w_strSql = w_strSql & " ,      KE.IDODATE "
            w_strSql = w_strSql & " ,      KE.SEQ "
            w_strSql = w_strSql & " ,      KE.ENDDATE "
            w_strSql = w_strSql & " ,      KE.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      KE.LASTUPDTIMEDATE "

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals((General.gInstall_Enum.AccessType_PassThrough).ToString) Then 'ORACLE

                w_strSql = w_strSql & " FROM   NS_KENMUINFO_F KE "
                w_strSql = w_strSql & " ,      NS_WARDDEPT_M  WM "
                w_strSql = w_strSql & " ,      NS_KINMUDEPT_M KM "
                w_strSql = w_strSql & " ,      NS_POST_M      PM "
                w_strSql = w_strSql & " WHERE  KE.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    KE.STAFFMNGID    = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    WM.HOSPITALCD  (+)  = KE.HOSPITALCD "
                w_strSql = w_strSql & " AND    WM.WARDDEPTCD  (+)  = KE.WARDDEPTCD "
                w_strSql = w_strSql & " AND    KM.HOSPITALCD  (+)  = KE.HOSPITALCD "
                w_strSql = w_strSql & " AND    KM.KINMUDEPTCD (+)  = KE.KINMUDEPTCD "
                w_strSql = w_strSql & " AND    PM.HOSPITALCD  (+)  = KE.HOSPITALCD "
                w_strSql = w_strSql & " AND    PM.POSTCD      (+)  = KE.POSTCD "

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_KENMUINFO_F KE "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_WARDDEPT_M  WM "
                w_strSql = w_strSql & " ON     WM.HOSPITALCD    = KE.HOSPITALCD "
                w_strSql = w_strSql & " AND    WM.WARDDEPTCD    = KE.WARDDEPTCD "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_KINMUDEPT_M KM "
                w_strSql = w_strSql & " ON     KM.HOSPITALCD    = KE.HOSPITALCD "
                w_strSql = w_strSql & " AND    KM.KINMUDEPTCD   = KE.KINMUDEPTCD "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_POST_M PM "
                w_strSql = w_strSql & " ON     PM.HOSPITALCD    = KE.HOSPITALCD "
                w_strSql = w_strSql & " AND    PM.POSTCD        = KE.POSTCD "
                w_strSql = w_strSql & " WHERE  KE.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    KE.STAFFMNGID    = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND KE.IDODATE      <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND   KE.IDODATE    <= " & m_numDateTo & " "
                w_strSql = w_strSql & " AND ( KE.ENDDATE >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " OR    KE.ENDDATE  = 0 "
                w_strSql = w_strSql & " OR    KE.ENDDATE IS NULL ) "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY KE.IDODATE ASC , KE.SEQ ASC"
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY KE.IDODATE DESC , KE.SEQ DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numKenmuKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numKenmuKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_KenmuIdo.Ido(m_numKenmuKensu)

                    w_WardDeptCD_F = .Fields("WARDDEPTCD")
                    w_WardDeptName_F = .Fields("WARDDEPTNAME")
                    w_KinmuDeptCD_F = .Fields("KINMUDEPTCD")
                    w_KinmuDeptName_F = .Fields("KINMUDEPTNAME")
                    w_PostCD_F = .Fields("POSTCD")
                    w_PostName_F = .Fields("POSTNAME")
                    w_SEQ_F = .Fields("SEQ")
                    w_DateFrom_F = .Fields("IDODATE")
                    w_DateTo_F = .Fields("ENDDATE")
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE")
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE")

                    g_KenmuIdo.strHospitalCD = m_strHospitalCD
                    g_KenmuIdo.strStaffMngID = m_strStaffMngID

                    For w_numLoop = 1 To m_numKenmuKensu
                        '�z�������R�[�h
                        g_KenmuIdo.Ido(w_numLoop).strWardDeptCD = General.paGetDbFieldVal(w_WardDeptCD_F, "")
                        '�z����������
                        g_KenmuIdo.Ido(w_numLoop).strWardDeptName = General.paGetDbFieldVal(w_WardDeptName_F, "")
                        '�Ζ������R�[�h
                        g_KenmuIdo.Ido(w_numLoop).strKinmuDeptCD = General.paGetDbFieldVal(w_KinmuDeptCD_F, "")
                        '�Ζ���������
                        g_KenmuIdo.Ido(w_numLoop).strKinmuDeptName = General.paGetDbFieldVal(w_KinmuDeptName_F, "")
                        '��E�R�[�h
                        g_KenmuIdo.Ido(w_numLoop).strPostCD = General.paGetDbFieldVal(w_PostCD_F, "")
                        '��E����
                        g_KenmuIdo.Ido(w_numLoop).strPostName = General.paGetDbFieldVal(w_PostName_F, "")
                        '�J�n��
                        g_KenmuIdo.Ido(w_numLoop).numDateFrom = Integer.Parse(General.paGetDbFieldVal(w_DateFrom_F, 0))
                        'SEQ
                        g_KenmuIdo.Ido(w_numLoop).numSEQ = Integer.Parse(General.paGetDbFieldVal(w_SEQ_F, 0))
                        '�I����
                        g_KenmuIdo.Ido(w_numLoop).numDateTo = Integer.Parse(General.paGetDbFieldVal(w_DateTo_F, 99999999))
                        '����o�^����
                        g_KenmuIdo.Ido(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_KenmuIdo.Ido(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetKenmuIdo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch ex As Exception
            Throw
        End Try
    End Function




    ''' <summary>
    ''' �Ƌ������擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetMenkyoInfo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetMenkyoInfo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Const W_strMenkyo As String = "S006" '�Ƌ����̗p�@�ėp�}�X�^ID


        Dim w_Rs As ADODB.Recordset
        Dim w_CD_F As ADODB.Field '�R�[�h
        Dim w_Name_F As ADODB.Field '����
        Dim w_No_F As ADODB.Field '�ԍ�

        '2012/10/25 fujisawa add st --------------
        Dim w_JapanAreaCD As ADODB.Field  '�s���{���R�[�h
        Dim w_JapanAreaName As ADODB.Field '�s���{������
        '2012/10/25 fujisawa add end --------------

        Dim w_GetDate_F As ADODB.Field '�擾�N����
        Dim w_Bikou_F As ADODB.Field '���l
        Dim w_FirstTime_F As ADODB.Field '����o�^����
        Dim w_UpdTime_F As ADODB.Field '�ŏI�X�V����

        fncGetMenkyoInfo = False
        Try
            ReDim g_MenkyoInfo(0)

            w_strSql = ""
            w_strSql = w_strSql & " SELECT LI.LICENSEKBN "

            '2012/10/25 fujisawa ch st ************************
            'w_strSql = w_strSql & " ,      HM.NAME "
            w_strSql = w_strSql & " ,      HM.NAME AS LICENSE "
            '2012/10/25 fujisawa ch end ***********************

            '2012/10/25 fujisawa add�@st -----------------------------
            '�s���{��
            w_strSql = w_strSql & " ,      HM2.MASTERCD "
            w_strSql = w_strSql & " ,      HM2.NAME AS JAPANAREANAME "
            '2012/10/25 fujisawa add end -----------------------------

            w_strSql = w_strSql & " ,      LI.LICENSENO "
            w_strSql = w_strSql & " ,      LI.GETDATE "
            w_strSql = w_strSql & " ,      LI.BIKOU "
            w_strSql = w_strSql & " ,      LI.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      LI.LASTUPDTIMEDATE "

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals((General.gInstall_Enum.AccessType_PassThrough).ToString) Then 'ORACLE

                'w_strSql = w_strSql & " FROM   NS_LICENSEINFO_F LI "
                'w_strSql = w_strSql & " ,      NS_HANYOU_M      HM "
                'w_strSql = w_strSql & " WHERE  LI.HOSPITALCD    = '" & m_strHospitalCD & "' "
                'w_strSql = w_strSql & " AND    LI.STAFFMNGID    = '" & m_strStaffMngID & "' "
                'w_strSql = w_strSql & " AND    HM.HOSPITALCD  (+)  = LI.HOSPITALCD "
                'w_strSql = w_strSql & " AND    HM.MASTERID    (+)  = '" & W_strMenkyo & "' "
                'w_strSql = w_strSql & " AND    HM.MASTERCD    (+)  = LI.LICENSEKBN "
                w_strSql = w_strSql & " FROM   NS_LICENSEINFO_F LI "
                w_strSql = w_strSql & " ,      NS_HANYOU_M      HM "
                w_strSql = w_strSql & " ,      NS_HANYOU_M      HM2 " '2012/10/25 fujisawa add
                w_strSql = w_strSql & " WHERE  LI.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    LI.STAFFMNGID    = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    HM.HOSPITALCD  (+)  = LI.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM.MASTERID    (+)  = '" & W_strMenkyo & "' "
                w_strSql = w_strSql & " AND    HM.MASTERCD    (+)  = LI.LICENSEKBN "

                '2012/10/25 fujisawa add st ---------------------------------------------------------------
                w_strSql = w_strSql & " AND    HM2.HOSPITALCD  (+)  = LI.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM2.MASTERID    (+)  = '" & General.G_MASTERID_JPAREACD & "' "
                w_strSql = w_strSql & " AND    HM2.MASTERCD    (+)  = LI.JAPANAREACD "
                '2012/10/25 fujisawa add end --------------------------------------------------------------

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_LICENSEINFO_F LI "

                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M HM "
                w_strSql = w_strSql & " ON     HM.HOSPITALCD    = LI.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM.MASTERID      = '" & W_strMenkyo & "' "
                w_strSql = w_strSql & " AND    HM.MASTERCD      = LI.LICENSEKBN "

                '2012/10/25 fujisawa add st ---------------------------------------------------------------
                '�s���{���ǉ�
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M HM2 "
                w_strSql = w_strSql & " ON     HM2.HOSPITALCD    = LI.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM2.MASTERID      = '" & General.G_MASTERID_JPAREACD & "' "
                w_strSql = w_strSql & " AND    HM2.MASTERCD      = LI.JAPANAREACD "
                '2012/10/25 fujisawa add end --------------------------------------------------------------

                w_strSql = w_strSql & " WHERE  LI.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    LI.STAFFMNGID    = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND LI.GETDATE      <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then


                w_strSql = w_strSql & " AND LI.GETDATE      >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " AND LI.GETDATE      <= " & m_numDateTo & " "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY LI.GETDATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY LI.GETDATE DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numMenkyoKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numMenkyoKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_MenkyoInfo(m_numMenkyoKensu)

                    w_CD_F = .Fields("LICENSEKBN")

                    '2012/10/25 fujisawa ch st **
                    'w_Name_F = .Fields("Name")
                    w_Name_F = .Fields("LICENSE")
                    '2012/10/25 fujisawa ch end *

                    w_No_F = .Fields("LICENSENO")

                    '2012/10/25 fujisawa add st -----------
                    '�s���{���R�[�h�E����
                    w_JapanAreaCD = .Fields("MASTERCD")
                    w_JapanAreaName = .Fields("JAPANAREANAME")
                    '2012/10/25 fujisawa add end ----------

                    w_GetDate_F = .Fields("GETDATE")
                    w_Bikou_F = .Fields("BIKOU")
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE")
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE")

                    For w_numLoop = 1 To m_numMenkyoKensu

                        g_MenkyoInfo(w_numLoop).strHospitalCD = m_strHospitalCD
                        g_MenkyoInfo(w_numLoop).strStaffMngID = m_strStaffMngID

                        '�Ƌ��R�[�h
                        g_MenkyoInfo(w_numLoop).strCD = General.paGetDbFieldVal(w_CD_F, "")
                        '�Ƌ�����
                        g_MenkyoInfo(w_numLoop).strName = General.paGetDbFieldVal(w_Name_F, "")
                        '�Ƌ��ԍ�
                        g_MenkyoInfo(w_numLoop).strNo = General.paGetDbFieldVal(w_No_F, "")

                        '2012/10/25 fujisawa add st ---------------------------------------
                        '�s���{���R�[�h
                        g_MenkyoInfo(w_numLoop).strJapanAreaCD = General.paGetDbFieldVal(w_JapanAreaCD, "")
                        '�s���{������
                        g_MenkyoInfo(w_numLoop).strJapanAreaName = General.paGetDbFieldVal(w_JapanAreaName, "")
                        '2012/10/25 fujisawa add end --------------------------------------

                        '�擾�N����
                        g_MenkyoInfo(w_numLoop).numGetDate = Integer.Parse(General.paGetDbFieldVal(w_GetDate_F, 0))
                        '���l
                        g_MenkyoInfo(w_numLoop).strBikou = General.paGetDbFieldVal(w_Bikou_F, "")
                        '����o�^����
                        g_MenkyoInfo(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_MenkyoInfo(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))


                        g_MenkyoInfo(w_numLoop).numDateFrom = 0
                        g_MenkyoInfo(w_numLoop).numDateTo = 0

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetMenkyoInfo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���i�����擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetShikakuInfo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetShikakuInfo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Const W_strShikaku As String = "S005" '���i���̗p�@�ėp�}�X�^ID

        Dim w_Rs As ADODB.Recordset
        Dim w_CD_F As ADODB.Field '�R�[�h
        Dim w_Name_F As ADODB.Field '����
        Dim w_GetDate_F As ADODB.Field '�擾�N����
        Dim w_DateFrom_F As ADODB.Field '�J�n�N����
        Dim w_DateTo_F As ADODB.Field '�I���N����
        Dim w_Bikou_F As ADODB.Field '���l
        Dim w_FirstTime_F As ADODB.Field '����o�^����
        Dim w_UpdTime_F As ADODB.Field '�ŏI�X�V����


        fncGetShikakuInfo = False
        Try
            ReDim g_ShikakuInfo(0)

            w_strSql = ""
            w_strSql = w_strSql & " SELECT QI.QUALIFYINGCD "
            w_strSql = w_strSql & " ,      HM.NAME "
            w_strSql = w_strSql & " ,      QI.GETDATE "
            w_strSql = w_strSql & " ,      QI.EFFFROMDATE "
            w_strSql = w_strSql & " ,      QI.EFFTODATE "
            w_strSql = w_strSql & " ,      QI.BIKOU "
            w_strSql = w_strSql & " ,      QI.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      QI.LASTUPDTIMEDATE "

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, CStr(General.gInstall_Enum.AccessType_PassThrough)).Equals(CStr(General.gInstall_Enum.AccessType_PassThrough)) Then 'ORACLE

                w_strSql = w_strSql & " FROM   NS_QUALIFYINFO_F QI "
                w_strSql = w_strSql & " ,      NS_HANYOU_M      HM "
                w_strSql = w_strSql & " WHERE  QI.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    QI.STAFFMNGID    = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    HM.HOSPITALCD  (+)  = QI.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM.MASTERID    (+)  = '" & W_strShikaku & "' "
                w_strSql = w_strSql & " AND    HM.MASTERCD    (+)  = QI.QUALIFYINGCD "

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_QUALIFYINFO_F QI "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M HM "
                w_strSql = w_strSql & " ON     HM.HOSPITALCD    = QI.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM.MASTERID      = '" & W_strShikaku & "' "
                w_strSql = w_strSql & " AND    HM.MASTERCD      = QI.QUALIFYINGCD "
                w_strSql = w_strSql & " WHERE  QI.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    QI.STAFFMNGID    = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND QI.EFFFROMDATE   <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND   QI.EFFFROMDATE <= " & m_numDateTo & " "
                w_strSql = w_strSql & " AND ( QI.EFFTODATE   >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " OR    QI.EFFTODATE    = 0 "
                w_strSql = w_strSql & " OR    QI.EFFTODATE   IS NULL ) "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY QI.EFFFROMDATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY QI.EFFFROMDATE DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numShikakuKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numShikakuKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_ShikakuInfo(m_numShikakuKensu)

                    w_CD_F = .Fields("QUALIFYINGCD")
                    w_Name_F = .Fields("NAME")
                    w_GetDate_F = .Fields("GETDATE")
                    w_DateFrom_F = .Fields("EFFFROMDATE")
                    w_DateTo_F = .Fields("EFFTODATE")
                    w_Bikou_F = .Fields("BIKOU")
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE")
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE")

                    For w_numLoop = 1 To m_numShikakuKensu

                        g_ShikakuInfo(w_numLoop).strHospitalCD = m_strHospitalCD
                        g_ShikakuInfo(w_numLoop).strStaffMngID = m_strStaffMngID

                        '���i�R�[�h
                        g_ShikakuInfo(w_numLoop).strCD = General.paGetDbFieldVal(w_CD_F, "")
                        '���i����
                        g_ShikakuInfo(w_numLoop).strName = General.paGetDbFieldVal(w_Name_F, "")
                        '�擾�N����
                        g_ShikakuInfo(w_numLoop).numGetDate = Integer.Parse(General.paGetDbFieldVal(w_GetDate_F, 0))
                        '�J�n�N����
                        g_ShikakuInfo(w_numLoop).numDateFrom = Integer.Parse(General.paGetDbFieldVal(w_DateFrom_F, 0))
                        '�I���N����
                        g_ShikakuInfo(w_numLoop).numDateTo = Integer.Parse(General.paGetDbFieldVal(w_DateTo_F, 99999999))
                        '���l
                        g_ShikakuInfo(w_numLoop).strBikou = General.paGetDbFieldVal(w_Bikou_F, "")
                        '����o�^����
                        g_ShikakuInfo(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_ShikakuInfo(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))

                        g_ShikakuInfo(w_numLoop).strNo = ""

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetShikakuInfo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �ψ������擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetIinInfo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetIinInfo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Const W_strIin As String = "S004" '�ψ����̗p�@�ėp�}�X�^ID
        Const W_strIinPost As String = G_MASTERID_IINPOSTNAME '�ψ���E���̗p�@�ėp�}�X�^ID

        Dim w_Rs As ADODB.Recordset
        Dim w_CD_F As ADODB.Field '�R�[�h
        Dim w_Name_F As ADODB.Field '����
        Dim w_DateFrom_F As ADODB.Field '�J�n�N����
        Dim w_DateTo_F As ADODB.Field '�I���N����
        Dim w_IinPostCd_F As ADODB.Field '�ψ���ECD
        Dim w_IinPostName_F As ADODB.Field '�ψ���E��
        Dim w_Bikou_F As ADODB.Field '���l
        Dim w_FirstTime_F As ADODB.Field '����o�^����
        Dim w_UpdTime_F As ADODB.Field '�ŏI�X�V����


        fncGetIinInfo = False

        ReDim g_IinInfo(0)
        Try
            w_strSql = ""
            w_strSql = w_strSql & " SELECT II.IINCD "
            w_strSql = w_strSql & " ,      HM.NAME "
            w_strSql = w_strSql & " ,      II.FROMDATE "
            w_strSql = w_strSql & " ,      II.TODATE "
            w_strSql = w_strSql & " ,      II.BIKOU "
            w_strSql = w_strSql & " ,      II.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      II.LASTUPDTIMEDATE "
            w_strSql = w_strSql & " ,      II.IINPOSTCD "
            w_strSql = w_strSql & " ,      HM_POST.NAME AS IINPOSTNAME "

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals((General.gInstall_Enum.AccessType_PassThrough).ToString) Then 'ORACLE

                w_strSql = w_strSql & " FROM   NS_IININFO_F II "
                w_strSql = w_strSql & " ,      NS_HANYOU_M  HM "
                w_strSql = w_strSql & " ,      NS_HANYOU_M  HM_POST "
                w_strSql = w_strSql & " WHERE  II.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    II.STAFFMNGID    = '" & m_strStaffMngID & "' "

                w_strSql = w_strSql & " AND    HM.HOSPITALCD  (+)  = II.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM.MASTERID    (+)  = '" & W_strIin & "' "
                w_strSql = w_strSql & " AND    HM.MASTERCD    (+)  = II.IINCD "

                w_strSql = w_strSql & " AND    HM_POST.HOSPITALCD  (+)  = II.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM_POST.MASTERID    (+)  = '" & W_strIinPost & "' "
                w_strSql = w_strSql & " AND    HM_POST.MASTERCD    (+)  = II.IINPOSTCD "

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_IININFO_F II "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M HM "
                w_strSql = w_strSql & " ON     HM.HOSPITALCD    = II.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM.MASTERID      = '" & W_strIin & "' "
                w_strSql = w_strSql & " AND    HM.MASTERCD      = II.IINCD "

                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M HM_POST "
                w_strSql = w_strSql & " ON     HM_POST.HOSPITALCD    = II.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM_POST.MASTERID      = '" & W_strIinPost & "' "
                w_strSql = w_strSql & " AND    HM_POST.MASTERCD      = II.IINPOSTCD "

                w_strSql = w_strSql & " WHERE  II.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    II.STAFFMNGID    = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND II.FROMDATE   <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND   II.FROMDATE <= " & m_numDateTo & " "
                w_strSql = w_strSql & " AND ( II.TODATE   >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " OR    II.TODATE    = 0 "
                w_strSql = w_strSql & " OR    II.TODATE   IS NULL ) "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY II.FROMDATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY II.FROMDATE DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numIinKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numIinKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_IinInfo(m_numIinKensu)

                    w_CD_F = .Fields("IINCD")
                    w_Name_F = .Fields("NAME")
                    w_DateFrom_F = .Fields("FROMDATE")
                    w_DateTo_F = .Fields("TODATE")
                    w_Bikou_F = .Fields("BIKOU")
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE")
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE")
                    w_IinPostCd_F = .Fields("IINPOSTCD")
                    w_IinPostName_F = .Fields("IINPOSTNAME")

                    For w_numLoop = 1 To m_numIinKensu

                        g_IinInfo(w_numLoop).strHospitalCD = m_strHospitalCD
                        g_IinInfo(w_numLoop).strStaffMngID = m_strStaffMngID

                        '�ψ��R�[�h
                        g_IinInfo(w_numLoop).strCD = General.paGetDbFieldVal(w_CD_F, "")
                        '�ψ�����
                        g_IinInfo(w_numLoop).strName = General.paGetDbFieldVal(w_Name_F, "")
                        '�J�n�N����
                        g_IinInfo(w_numLoop).numDateFrom = Integer.Parse(General.paGetDbFieldVal(w_DateFrom_F, 0))
                        '�I���N����
                        g_IinInfo(w_numLoop).numDateTo = Integer.Parse(General.paGetDbFieldVal(w_DateTo_F, 99999999))
                        '���l
                        g_IinInfo(w_numLoop).strBikou = General.paGetDbFieldVal(w_Bikou_F, "")
                        '����o�^����
                        g_IinInfo(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_IinInfo(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))

                        '��E
                        g_IinInfo(w_numLoop).strIinPostCd = General.paGetDbFieldVal(w_IinPostCd_F, "")
                        g_IinInfo(w_numLoop).strIinPostName = General.paGetDbFieldVal(w_IinPostName_F, "")

                        g_IinInfo(w_numLoop).numGetDate = 0
                        g_IinInfo(w_numLoop).strNo = ""

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetIinInfo = True

            General.g_ErrorProc = w_strPreErrorProc
        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' �E�������擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetSyokurekiInfo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetSyokurekiInfo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Const W_strSyokureki As String = "S002" '�n��E�s���{���@�ėp�}�X�^ID

        Dim w_Rs As ADODB.Recordset
        Dim w_CD_F As ADODB.Field '�R�[�h
        Dim w_Name_F As ADODB.Field '����
        Dim w_DateFrom_F As ADODB.Field '�J�n�N����
        Dim w_DateTo_F As ADODB.Field '�I���N����
        Dim w_Area_F As ADODB.Field '�����@�֖�
        Dim w_ExpMedicalName_F As ADODB.Field '�o���f�É�
        Dim w_Bikou_F As ADODB.Field '���l
        Dim w_FirstTime_F As ADODB.Field '����o�^����
        Dim w_UpdTime_F As ADODB.Field '�ŏI�X�V����


        fncGetSyokurekiInfo = False
        Try
            ReDim g_SyokurekiInfo(0)

            w_strSql = ""
            w_strSql = w_strSql & " SELECT JC.JAPANAREACD "
            w_strSql = w_strSql & " ,      HM.NAME "
            w_strSql = w_strSql & " ,      JC.FROMDATE "
            w_strSql = w_strSql & " ,      JC.TODATE "
            w_strSql = w_strSql & " ,      JC.BELONGORGNAME "
            w_strSql = w_strSql & " ,      JC.EXPMEDICALNAME "
            w_strSql = w_strSql & " ,      JC.BIKOU "
            w_strSql = w_strSql & " ,      JC.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      JC.LASTUPDTIMEDATE "

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals((General.gInstall_Enum.AccessType_PassThrough).ToString) Then 'ORACLE

                w_strSql = w_strSql & " FROM   NS_JOBCAREERINFO_F JC "
                w_strSql = w_strSql & " ,      NS_HANYOU_M  HM "
                w_strSql = w_strSql & " WHERE  JC.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    JC.STAFFMNGID    = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    HM.HOSPITALCD  (+)  = JC.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM.MASTERID    (+)  = '" & W_strSyokureki & "' "
                w_strSql = w_strSql & " AND    HM.MASTERCD    (+)  = JC.JAPANAREACD "

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_JOBCAREERINFO_F JC "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M HM "
                w_strSql = w_strSql & " ON     HM.HOSPITALCD    = JC.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM.MASTERID      = '" & W_strSyokureki & "' "
                w_strSql = w_strSql & " AND    HM.MASTERCD      = JC.JAPANAREACD "
                w_strSql = w_strSql & " WHERE  JC.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    JC.STAFFMNGID    = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND JC.FROMDATE   <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND   JC.FROMDATE <= " & m_numDateTo & " "
                w_strSql = w_strSql & " AND ( JC.TODATE   >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " OR    JC.TODATE    = 0 "
                w_strSql = w_strSql & " OR    JC.TODATE   IS NULL ) "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY JC.FROMDATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY JC.FROMDATE DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numSyokurekiKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numSyokurekiKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_SyokurekiInfo(m_numSyokurekiKensu)

                    w_CD_F = .Fields("JAPANAREACD")
                    w_Name_F = .Fields("NAME")
                    w_DateFrom_F = .Fields("FROMDATE")
                    w_DateTo_F = .Fields("TODATE")
                    w_Area_F = .Fields("BELONGORGNAME")
                    w_ExpMedicalName_F = .Fields("EXPMEDICALNAME")
                    w_Bikou_F = .Fields("BIKOU")
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE")
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE")

                    For w_numLoop = 1 To m_numSyokurekiKensu

                        g_SyokurekiInfo(w_numLoop).strHospitalCD = m_strHospitalCD
                        g_SyokurekiInfo(w_numLoop).strStaffMngID = m_strStaffMngID
                        '�s���{���R�[�h
                        g_SyokurekiInfo(w_numLoop).strCD = General.paGetDbFieldVal(w_CD_F, "")
                        '�s���{������
                        g_SyokurekiInfo(w_numLoop).strName = General.paGetDbFieldVal(w_Name_F, "")
                        '�J�n�N����
                        g_SyokurekiInfo(w_numLoop).numDateFrom = Integer.Parse(General.paGetDbFieldVal(w_DateFrom_F, 0))
                        '�I���N����
                        g_SyokurekiInfo(w_numLoop).numDateTo = Integer.Parse(General.paGetDbFieldVal(w_DateTo_F, 99999999))
                        '�Ζ��@�֖�
                        g_SyokurekiInfo(w_numLoop).strArea = General.paGetDbFieldVal(w_Area_F, "")
                        '�o���f�É�'
                        g_SyokurekiInfo(w_numLoop).strExpMedicalName = General.paGetDbFieldVal(w_ExpMedicalName_F, "")
                        '���l
                        g_SyokurekiInfo(w_numLoop).strBikou = General.paGetDbFieldVal(w_Bikou_F, "")
                        '����o�^����
                        g_SyokurekiInfo(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_SyokurekiInfo(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetSyokurekiInfo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��ʊw�������擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetIppanGakurekiInfo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetIppanGakurekiInfo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Const W_strChiiki As String = "S002"
        Const W_strIppanGakureki As String = "S008" '�w�Z�敪���̗p�@�ėp�}�X�^ID
        Const W_strSchoolName As String = "S009"

        Dim w_Rs As ADODB.Recordset
        Dim w_Kbn_F As ADODB.Field '�R�[�h
        Dim w_KbnName_F As ADODB.Field '����
        Dim w_ChiikiCD_F As ADODB.Field '�n��R�[�h
        Dim w_ChiikiName_F As ADODB.Field '�n�於��
        Dim w_LastKbn_F As ADODB.Field '�ŏI�w���敪
        Dim w_LastDate_F As ADODB.Field '���ƔN����
        Dim w_SchoolCD_F As ADODB.Field '�w�Z�R�[�h
        Dim w_SchoolName_F As ADODB.Field '�w�Z��
        Dim w_Bikou_F As ADODB.Field '�C���ߒ�
        Dim w_FirstTime_F As ADODB.Field '����o�^����
        Dim w_UpdTime_F As ADODB.Field '�ŏI�X�V����


        fncGetIppanGakurekiInfo = False
        Try
            ReDim g_IppanGakurekiInfo(0)

            w_strSql = ""
            w_strSql = w_strSql & " SELECT GS.SCHHISKBN "
            w_strSql = w_strSql & " ,      H1.NAME "
            w_strSql = w_strSql & " ,      GS.AREACD "
            w_strSql = w_strSql & " ,      H2.NAME AS AREANAME "
            w_strSql = w_strSql & " ,      GS.FINALSCHHISKBN "
            w_strSql = w_strSql & " ,      GS.GRADUATEDATE "
            w_strSql = w_strSql & " ,      GS.SCHOOLNAMECD "
            w_strSql = w_strSql & " ,      H3.NAME AS SCHOOLNAME "
            w_strSql = w_strSql & " ,      GS.ENDCOURSE "
            w_strSql = w_strSql & " ,      GS.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      GS.LASTUPDTIMEDATE "

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals((General.gInstall_Enum.AccessType_PassThrough).ToString) Then 'ORACLE

                w_strSql = w_strSql & " FROM   NS_GENESCHHISINFO_F GS "
                w_strSql = w_strSql & " ,      NS_HANYOU_M  H1 "
                w_strSql = w_strSql & " ,      NS_HANYOU_M  H2 "
                w_strSql = w_strSql & " ,      NS_HANYOU_M  H3 "
                w_strSql = w_strSql & " WHERE  GS.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    GS.STAFFMNGID    = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    H1.HOSPITALCD  (+)  = GS.HOSPITALCD "
                w_strSql = w_strSql & " AND    H1.MASTERID    (+)  = '" & W_strIppanGakureki & "' "
                w_strSql = w_strSql & " AND    H1.MASTERCD    (+)  = GS.SCHHISKBN "
                w_strSql = w_strSql & " AND    H2.HOSPITALCD  (+)  = GS.HOSPITALCD "
                w_strSql = w_strSql & " AND    H2.MASTERID    (+)  = '" & W_strChiiki & "' "
                w_strSql = w_strSql & " AND    H2.MASTERCD    (+)  = GS.AREACD "
                w_strSql = w_strSql & " AND    H3.HOSPITALCD  (+)  = GS.HOSPITALCD "
                w_strSql = w_strSql & " AND    H3.MASTERID    (+)  = '" & W_strSchoolName & "' "
                w_strSql = w_strSql & " AND    H3.MASTERCD    (+)  = GS.SCHOOLNAMECD "

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_GENESCHHISINFO_F GS "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M  H1 "
                w_strSql = w_strSql & " ON     H1.HOSPITALCD    = GS.HOSPITALCD "
                w_strSql = w_strSql & " AND    H1.MASTERID      = '" & W_strIppanGakureki & "' "
                w_strSql = w_strSql & " AND    H1.MASTERCD      = GS.SCHHISKBN "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M  H2 "
                w_strSql = w_strSql & " ON     H2.HOSPITALCD    = GS.HOSPITALCD "
                w_strSql = w_strSql & " AND    H2.MASTERID      = '" & W_strChiiki & "' "
                w_strSql = w_strSql & " AND    H2.MASTERCD      = GS.AREACD "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M  H3 "
                w_strSql = w_strSql & " ON     H3.HOSPITALCD    = GS.HOSPITALCD "
                w_strSql = w_strSql & " AND    H3.MASTERID      = '" & W_strSchoolName & "' "
                w_strSql = w_strSql & " AND    H3.MASTERCD      = GS.SCHOOLNAMECD "
                w_strSql = w_strSql & " WHERE  GS.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    GS.STAFFMNGID    = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND GS.GRADUATEDATE   <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND GS.GRADUATEDATE   >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " AND GS.GRADUATEDATE   <= " & m_numDateTo & " "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY GS.GRADUATEDATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY GS.GRADUATEDATE DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numIppanGakurekiKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numIppanGakurekiKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_IppanGakurekiInfo(m_numIppanGakurekiKensu)

                    w_Kbn_F = .Fields("SCHHISKBN") '�R�[�h
                    w_KbnName_F = .Fields("NAME") '����
                    w_ChiikiCD_F = .Fields("AREACD") '�n��R�[�h
                    w_ChiikiName_F = .Fields("AREANAME") '�n�於��
                    w_LastKbn_F = .Fields("FINALSCHHISKBN") '�ŏI�w���敪
                    w_LastDate_F = .Fields("GRADUATEDATE") '���ƔN����
                    w_SchoolCD_F = .Fields("SCHOOLNAMECD") '�w�Z�R�[�h
                    w_SchoolName_F = .Fields("SCHOOLNAME") '�w�Z��
                    w_Bikou_F = .Fields("ENDCOURSE") '�C���ߒ�
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE") '����o�^����
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE") '�ŏI�X�V����

                    For w_numLoop = 1 To m_numIppanGakurekiKensu

                        g_IppanGakurekiInfo(w_numLoop).strHospitalCD = m_strHospitalCD
                        g_IppanGakurekiInfo(w_numLoop).strStaffMngID = m_strStaffMngID
                        '�敪
                        g_IppanGakurekiInfo(w_numLoop).strKbn = General.paGetDbFieldVal(w_Kbn_F, "")
                        '�敪����
                        g_IppanGakurekiInfo(w_numLoop).strKbnName = General.paGetDbFieldVal(w_KbnName_F, "")
                        '�n��R�[�h
                        g_IppanGakurekiInfo(w_numLoop).strChiikiCD = General.paGetDbFieldVal(w_ChiikiCD_F, "")
                        '�n�於��
                        g_IppanGakurekiInfo(w_numLoop).strChiikiName = General.paGetDbFieldVal(w_ChiikiName_F, "")
                        '�ŏI�w���敪
                        g_IppanGakurekiInfo(w_numLoop).strLastKbn = General.paGetDbFieldVal(w_LastKbn_F, "")
                        '���ƔN����
                        g_IppanGakurekiInfo(w_numLoop).numDate = Integer.Parse(General.paGetDbFieldVal(w_LastDate_F, 0))
                        '�w�Z�R�[�h
                        g_IppanGakurekiInfo(w_numLoop).strSchoolCD = General.paGetDbFieldVal(w_SchoolCD_F, "")
                        '�w�Z��
                        g_IppanGakurekiInfo(w_numLoop).strSchoolName = General.paGetDbFieldVal(w_SchoolName_F, "")
                        '�C���ߒ�
                        g_IppanGakurekiInfo(w_numLoop).strBikou = General.paGetDbFieldVal(w_Bikou_F, "")
                        '����o�^����
                        g_IppanGakurekiInfo(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_IppanGakurekiInfo(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))


                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetIppanGakurekiInfo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch ex As Exception
            Throw
        End Try
    End Function



    ''' <summary>
    ''' ���w�������擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetSenmonGakurekiInfo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetSenmonGakurekiInfo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Const W_strChiiki As String = "S002"
        Const W_strSenmonGakureki As String = "S010" '�w�Z�敪���̗p�@�ėp�}�X�^ID
        Const W_strSchoolName As String = "S011"

        Dim w_Rs As ADODB.Recordset
        Dim w_Kbn_F As ADODB.Field '�R�[�h
        Dim w_KbnName_F As ADODB.Field '����
        Dim w_ChiikiCD_F As ADODB.Field '�n��R�[�h
        Dim w_ChiikiName_F As ADODB.Field '�n�於��
        Dim w_LastKbn_F As ADODB.Field '�ŏI�w���敪
        Dim w_LastDate_F As ADODB.Field '���ƔN����
        Dim w_SchoolCD_F As ADODB.Field '�w�Z�R�[�h
        Dim w_SchoolName_F As ADODB.Field '�w�Z��
        Dim w_Bikou_F As ADODB.Field '�C���ߒ�
        Dim w_FirstTime_F As ADODB.Field '����o�^����
        Dim w_UpdTime_F As ADODB.Field '�ŏI�X�V����


        fncGetSenmonGakurekiInfo = False

        ReDim g_SenmonGakurekiInfo(0)
        Try
            w_strSql = ""
            w_strSql = w_strSql & " SELECT SS.SCHOOLKBN "
            w_strSql = w_strSql & " ,      H1.NAME "
            w_strSql = w_strSql & " ,      SS.AREACD "
            w_strSql = w_strSql & " ,      H2.NAME AS AREANAME "
            w_strSql = w_strSql & " ,      SS.FINALSCHHISKBN "
            w_strSql = w_strSql & " ,      SS.GRADUATEDATE "
            w_strSql = w_strSql & " ,      SS.SCHOOLNAMECD "
            w_strSql = w_strSql & " ,      H3.NAME AS SCHOOLNAME "
            w_strSql = w_strSql & " ,      SS.ENDCOURSE "
            w_strSql = w_strSql & " ,      SS.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      SS.LASTUPDTIMEDATE "

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals((General.gInstall_Enum.AccessType_PassThrough).ToString) Then 'ORACLE

                w_strSql = w_strSql & " FROM   NS_SPECSCHHISINFO_F SS "
                w_strSql = w_strSql & " ,      NS_HANYOU_M  H1 "
                w_strSql = w_strSql & " ,      NS_HANYOU_M  H2 "
                w_strSql = w_strSql & " ,      NS_HANYOU_M  H3 "
                w_strSql = w_strSql & " WHERE  SS.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    SS.STAFFMNGID    = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    H1.HOSPITALCD  (+)  = SS.HOSPITALCD "
                w_strSql = w_strSql & " AND    H1.MASTERID    (+)  = '" & W_strSenmonGakureki & "' "
                w_strSql = w_strSql & " AND    H1.MASTERCD    (+)  = SS.SCHOOLKBN "
                w_strSql = w_strSql & " AND    H2.HOSPITALCD  (+)  = SS.HOSPITALCD "
                w_strSql = w_strSql & " AND    H2.MASTERID    (+)  = '" & W_strChiiki & "' "
                w_strSql = w_strSql & " AND    H2.MASTERCD    (+)  = SS.AREACD "
                w_strSql = w_strSql & " AND    H3.HOSPITALCD  (+)  = SS.HOSPITALCD "
                w_strSql = w_strSql & " AND    H3.MASTERID    (+)  = '" & W_strSchoolName & "' "
                w_strSql = w_strSql & " AND    H3.MASTERCD    (+)  = SS.SCHOOLNAMECD "

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_SPECSCHHISINFO_F SS "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M H1 "
                w_strSql = w_strSql & " ON     H1.HOSPITALCD    = SS.HOSPITALCD "
                w_strSql = w_strSql & " AND    H1.MASTERID      = '" & W_strSenmonGakureki & "' "
                w_strSql = w_strSql & " AND    H1.MASTERCD      = SS.SCHOOLKBN "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M H2 "
                w_strSql = w_strSql & " ON     H2.HOSPITALCD    = SS.HOSPITALCD "
                w_strSql = w_strSql & " AND    H2.MASTERID      = '" & W_strChiiki & "' "
                w_strSql = w_strSql & " AND    H2.MASTERCD      = SS.AREACD "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M H3 "
                w_strSql = w_strSql & " ON     H3.HOSPITALCD    = SS.HOSPITALCD "
                w_strSql = w_strSql & " AND    H3.MASTERID      = '" & W_strSchoolName & "' "
                w_strSql = w_strSql & " AND    H3.MASTERCD      = SS.SCHOOLNAMECD "
                w_strSql = w_strSql & " WHERE  SS.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    SS.STAFFMNGID    = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND SS.GRADUATEDATE   <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND SS.GRADUATEDATE   >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " AND SS.GRADUATEDATE   <= " & m_numDateTo & " "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY SS.GRADUATEDATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY SS.GRADUATEDATE DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numSenmonGakurekiKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numSenmonGakurekiKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_SenmonGakurekiInfo(m_numSenmonGakurekiKensu)

                    w_Kbn_F = .Fields("SCHOOLKBN") '�R�[�h
                    w_KbnName_F = .Fields("NAME") '����
                    w_ChiikiCD_F = .Fields("AREACD") '�n��R�[�h
                    w_ChiikiName_F = .Fields("AREANAME") '�n�於��
                    w_LastKbn_F = .Fields("FINALSCHHISKBN") '�ŏI�w���敪
                    w_LastDate_F = .Fields("GRADUATEDATE") '���ƔN����
                    w_SchoolCD_F = .Fields("SCHOOLNAMECD") '�w�Z�R�[�h
                    w_SchoolName_F = .Fields("SCHOOLNAME") '�w�Z��
                    w_Bikou_F = .Fields("ENDCOURSE") '�C���ߒ�
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE") '����o�^����
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE") '�ŏI�X�V����

                    For w_numLoop = 1 To m_numSenmonGakurekiKensu

                        g_SenmonGakurekiInfo(w_numLoop).strHospitalCD = m_strHospitalCD
                        g_SenmonGakurekiInfo(w_numLoop).strStaffMngID = m_strStaffMngID

                        '�敪
                        g_SenmonGakurekiInfo(w_numLoop).strKbn = General.paGetDbFieldVal(w_Kbn_F, "")
                        '�敪����
                        g_SenmonGakurekiInfo(w_numLoop).strKbnName = General.paGetDbFieldVal(w_KbnName_F, "")
                        '�n��R�[�h
                        g_SenmonGakurekiInfo(w_numLoop).strChiikiCD = General.paGetDbFieldVal(w_ChiikiCD_F, "")
                        '�n�於��
                        g_SenmonGakurekiInfo(w_numLoop).strChiikiName = General.paGetDbFieldVal(w_ChiikiName_F, "")
                        '�ŏI�w���敪
                        g_SenmonGakurekiInfo(w_numLoop).strLastKbn = General.paGetDbFieldVal(w_LastKbn_F, "")
                        '���ƔN����
                        g_SenmonGakurekiInfo(w_numLoop).numDate = Integer.Parse(General.paGetDbFieldVal(w_LastDate_F, 0))
                        '�w�Z�R�[�h
                        g_SenmonGakurekiInfo(w_numLoop).strSchoolCD = General.paGetDbFieldVal(w_SchoolCD_F, "")
                        '�w�Z��
                        g_SenmonGakurekiInfo(w_numLoop).strSchoolName = General.paGetDbFieldVal(w_SchoolName_F, "")
                        '�C���ߒ�
                        g_SenmonGakurekiInfo(w_numLoop).strBikou = General.paGetDbFieldVal(w_Bikou_F, "")
                        '����o�^����
                        g_SenmonGakurekiInfo(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_SenmonGakurekiInfo(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetSenmonGakurekiInfo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch ex As Exception
            Throw
        End Try
    End Function



    ''' <summary>
    ''' ���x�����擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetChoukyuInfo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetChoukyuInfo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Const W_strChoukyu As String = "S012" '���x���̗p�@�ėp�}�X�^ID

        Dim w_Rs As ADODB.Recordset
        Dim w_CD_F As ADODB.Field '�R�[�h
        Dim w_Name_F As ADODB.Field '����
        Dim w_SecName_F As ADODB.Field '2018/10/02 Darren ADD
        Dim w_DateFrom_F As ADODB.Field '�J�n�N����
        Dim w_DateTo_F As ADODB.Field '�I���N����
        Dim w_Bikou_F As ADODB.Field '���l
        Dim w_WeeklyTime_F As ADODB.Field '�T�J������ '2018/08/24 T.K add
        Dim w_FirstTime_F As ADODB.Field '����o�^����
        Dim w_UpdTime_F As ADODB.Field '�ŏI�X�V����


        fncGetChoukyuInfo = False
        Try
            ReDim g_ChoukyuInfo(0)

            w_strSql = ""
            w_strSql = w_strSql & " SELECT LL.HOLIDAYCD "
            w_strSql = w_strSql & " ,      HM.NAME "
            w_strSql = w_strSql & " ,      HM.SECNAME "
            w_strSql = w_strSql & " ,      LL.FROMDATE "
            w_strSql = w_strSql & " ,      LL.TODATE "
            w_strSql = w_strSql & " ,      LL.BIKOU "
            w_strSql = w_strSql & " ,      LL.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      LL.LASTUPDTIMEDATE "
            w_strSql = w_strSql & " ,      LL.WEEKLYTIME " '2018/08/24 T.K add

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals((General.gInstall_Enum.AccessType_PassThrough).ToString) Then 'ORACLE

                w_strSql = w_strSql & " FROM   NS_LONGLEAVEINFO_F LL "
                w_strSql = w_strSql & " ,      NS_HANYOU_M        HM "
                w_strSql = w_strSql & " WHERE  LL.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    LL.STAFFMNGID    = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    HM.HOSPITALCD  (+)  = LL.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM.MASTERID    (+)  = '" & W_strChoukyu & "' "
                w_strSql = w_strSql & " AND    HM.MASTERCD    (+)  = LL.HOLIDAYCD "

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_LONGLEAVEINFO_F LL "
                w_strSql = w_strSql & " LEFT OUTER JOIN  NS_HANYOU_M        HM "
                w_strSql = w_strSql & " ON     HM.HOSPITALCD    = LL.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM.MASTERID      = '" & W_strChoukyu & "' "
                w_strSql = w_strSql & " AND    HM.MASTERCD      = LL.HOLIDAYCD "
                w_strSql = w_strSql & " WHERE  LL.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    LL.STAFFMNGID    = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND LL.FROMDATE   <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND   LL.FROMDATE <= " & m_numDateTo & " "
                w_strSql = w_strSql & " AND ( LL.TODATE   >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " OR    LL.TODATE    = 0 "
                w_strSql = w_strSql & " OR    LL.TODATE   IS NULL ) "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY LL.FROMDATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY LL.FROMDATE DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numChoukyuKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numChoukyuKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_ChoukyuInfo(m_numChoukyuKensu)

                    w_CD_F = .Fields("HOLIDAYCD") '�R�[�h
                    w_Name_F = .Fields("NAME") '����
                    w_SecName_F = .Fields("SECNAME") '2018/10/02 Darren ADD
                    w_DateFrom_F = .Fields("FROMDATE") '�J�n�N����
                    w_DateTo_F = .Fields("TODATE") '�I���N����
                    w_Bikou_F = .Fields("BIKOU") '���l
                    w_WeeklyTime_F = .Fields("WEEKLYTIME") '�T�J������ '2018/08/24 T.K add
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE") '����o�^����
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE") '�ŏI�X�V����

                    For w_numLoop = 1 To m_numChoukyuKensu

                        g_ChoukyuInfo(w_numLoop).strHospitalCD = m_strHospitalCD
                        g_ChoukyuInfo(w_numLoop).strStaffMngID = m_strStaffMngID

                        '�R�[�h
                        g_ChoukyuInfo(w_numLoop).strCD = General.paGetDbFieldVal(w_CD_F, "")
                        '����
                        g_ChoukyuInfo(w_numLoop).strName = General.paGetDbFieldVal(w_Name_F, "")
                        '���O�̗���
                        g_ChoukyuInfo(w_numLoop).strSecName = General.paGetDbFieldVal(w_SecName_F, "") '2018/10/02 Darren ADD
                        '�J�n�N����
                        g_ChoukyuInfo(w_numLoop).numDateFrom = Integer.Parse(General.paGetDbFieldVal(w_DateFrom_F, 0))
                        '�I���N����
                        g_ChoukyuInfo(w_numLoop).numDateTo = Integer.Parse(General.paGetDbFieldVal(w_DateTo_F, 99999999))
                        '���l
                        g_ChoukyuInfo(w_numLoop).strBikou = General.paGetDbFieldVal(w_Bikou_F, "")
                        '2018/08/24 T.K add st -------------------------------------
                        '�T�J������
                        g_ChoukyuInfo(w_numLoop).numWeeklyTime = Integer.Parse(General.paGetDbFieldVal(w_WeeklyTime_F, 0))
                        '2018/08/24 T.K add ed -------------------------------------
                        '����o�^����
                        g_ChoukyuInfo(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_ChoukyuInfo(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))

                        g_ChoukyuInfo(w_numLoop).numGetDate = 0
                        g_ChoukyuInfo(w_numLoop).strNo = ""

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetChoukyuInfo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Y�x�����擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetSankyuInfo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetSankyuInfo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Dim w_Rs As ADODB.Recordset
        Dim w_PlanDate_F As ADODB.Field '�\��N������
        Dim w_TwinFlg_F As ADODB.Field '�o�ً敪
        Dim w_BirthDate_F As ADODB.Field '�o�Y�N����
        Dim w_PlanSanzenYamenFrom_F As ADODB.Field '�\��Y�O���From
        Dim w_PlanSanzenYamenTo_F As ADODB.Field '�\��Y�O���To
        Dim w_PlanSanzenHolFrom_F As ADODB.Field '�\��Y�O�x��From
        Dim w_PlanSanzenHolTo_F As ADODB.Field '�\��Y�O�x��To
        Dim w_PlanSangoHolFrom_F As ADODB.Field '�\��Y��x��From
        Dim w_PlanSangoHolTo_F As ADODB.Field '�\��Y��x��To
        Dim w_PlanIkujiHolFrom_F As ADODB.Field '�\��玙�x��From
        Dim w_PlanIkujiHolTo_F As ADODB.Field '�\��玙�x��To
        Dim w_FixedSanzenYamenFrom_F As ADODB.Field '�m��Y�O���From
        Dim w_FixedSanzenYamenTo_F As ADODB.Field '�m��Y�O���To
        Dim w_FixedSanzenHolFrom_F As ADODB.Field '�m��Y�O�x��From
        Dim w_FixedSanzenHolTo_F As ADODB.Field '�m��Y�O�x��To
        Dim w_FixedSangoHolFrom_F As ADODB.Field '�m��Y��x��From
        Dim w_FixedSangoHolTo_F As ADODB.Field '�m��Y��x��To
        Dim w_FixedIkujiHolFrom_F As ADODB.Field '�m��玙�x��From
        Dim w_FixedIkujiHolTo_F As ADODB.Field '�m��玙�x��To
        Dim w_SEQ_F As ADODB.Field 'SEQ
        Dim w_FirstTime_F As ADODB.Field '����o�^����
        Dim w_UpdTime_F As ADODB.Field '�ŏI�X�V����


        fncGetSankyuInfo = False
        Try
            ReDim g_SankyuInfo(0)

            w_strSql = ""
            w_strSql = w_strSql & " SELECT SK.PLANDATE "
            w_strSql = w_strSql & " ,      SK.TWINFLG "
            w_strSql = w_strSql & " ,      SK.BIRTHDATE "
            w_strSql = w_strSql & " ,      SK.PLANSANZENYAMENFM "
            w_strSql = w_strSql & " ,      SK.PLANSANZENYAMENTO "
            w_strSql = w_strSql & " ,      SK.PLANSANZENHOLFM "
            w_strSql = w_strSql & " ,      SK.PLANSANZENHOLTO "
            w_strSql = w_strSql & " ,      SK.PLANSANGOHOLFM "
            w_strSql = w_strSql & " ,      SK.PLANSANGOHOLTO "
            w_strSql = w_strSql & " ,      SK.PLANIKUJIHOLFM "
            w_strSql = w_strSql & " ,      SK.PLANIKUJIHOLTO "
            w_strSql = w_strSql & " ,      SK.FIXEDSANZENYAMENFM "
            w_strSql = w_strSql & " ,      SK.FIXEDSANZENYAMENTO "
            w_strSql = w_strSql & " ,      SK.FIXEDSANZENHOLFM "
            w_strSql = w_strSql & " ,      SK.FIXEDSANZENHOLTO "
            w_strSql = w_strSql & " ,      SK.FIXEDSANGOHOLFM "
            w_strSql = w_strSql & " ,      SK.FIXEDSANGOHOLTO "
            w_strSql = w_strSql & " ,      SK.FIXEDIKUJIHOLFM "
            w_strSql = w_strSql & " ,      SK.FIXEDIKUJIHOLTO "
            w_strSql = w_strSql & " ,      SK.SEQ "
            w_strSql = w_strSql & " ,      SK.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      SK.LASTUPDTIMEDATE "
            w_strSql = w_strSql & " FROM   NS_SANKYUINFO_F SK "
            w_strSql = w_strSql & " WHERE  SK.HOSPITALCD    = '" & m_strHospitalCD & "' "
            w_strSql = w_strSql & " AND    SK.STAFFMNGID    = '" & m_strStaffMngID & "' "
            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND SK.PLANDATE   <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND SK.PLANDATE   >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " AND SK.PLANDATE   <= " & m_numDateTo & " "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY SK.PLANDATE ASC , SK.SEQ ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY SK.PLANDATE DESC , SK.SEQ DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numSankyuKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numSankyuKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_SankyuInfo(m_numSankyuKensu)

                    w_PlanDate_F = .Fields("PLANDATE") '�\��N������
                    w_TwinFlg_F = .Fields("TWINFLG") '�o�ً敪
                    w_BirthDate_F = .Fields("BIRTHDATE") '�o�Y�N����
                    w_PlanSanzenYamenFrom_F = .Fields("PLANSANZENYAMENFM") '�\��Y�O���From
                    w_PlanSanzenYamenTo_F = .Fields("PLANSANZENYAMENTO") '�\��Y�O���To
                    w_PlanSanzenHolFrom_F = .Fields("PLANSANZENHOLFM") '�\��Y�O�x��From
                    w_PlanSanzenHolTo_F = .Fields("PLANSANZENHOLTO") '�\��Y�O�x��To
                    w_PlanSangoHolFrom_F = .Fields("PLANSANGOHOLFM") '�\��Y��x��From
                    w_PlanSangoHolTo_F = .Fields("PLANSANGOHOLTO") '�\��Y��x��To
                    w_PlanIkujiHolFrom_F = .Fields("PLANIKUJIHOLFM") '�\��玙�x��From
                    w_PlanIkujiHolTo_F = .Fields("PLANIKUJIHOLTO") '�\��玙�x��To
                    w_FixedSanzenYamenFrom_F = .Fields("FIXEDSANZENYAMENFM") '�m��Y�O���From
                    w_FixedSanzenYamenTo_F = .Fields("FIXEDSANZENYAMENTO") '�m��Y�O���To
                    w_FixedSanzenHolFrom_F = .Fields("FIXEDSANZENHOLFM") '�m��Y�O�x��From
                    w_FixedSanzenHolTo_F = .Fields("FIXEDSANZENHOLTO") '�m��Y�O�x��To
                    w_FixedSangoHolFrom_F = .Fields("FIXEDSANGOHOLFM") '�m��Y��x��From
                    w_FixedSangoHolTo_F = .Fields("FIXEDSANGOHOLTO") '�m��Y��x��To
                    w_FixedIkujiHolFrom_F = .Fields("FIXEDIKUJIHOLFM") '�m��玙�x��From
                    w_FixedIkujiHolTo_F = .Fields("FIXEDIKUJIHOLTO") '�m��玙�x��To
                    w_SEQ_F = .Fields("SEQ") 'SEQ
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE") '����o�^����
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE") '�ŏI�X�V����

                    For w_numLoop = 1 To m_numSankyuKensu

                        g_SankyuInfo(w_numLoop).strHospitalCD = m_strHospitalCD
                        g_SankyuInfo(w_numLoop).strStaffMngID = m_strStaffMngID

                        'SEQ
                        g_SankyuInfo(w_numLoop).numSEQ = General.paGetDbFieldVal(w_SEQ_F, 0)
                        '�\��N������
                        g_SankyuInfo(w_numLoop).numPlanDate = Integer.Parse(General.paGetDbFieldVal(w_PlanDate_F, 0))
                        '�o�ً敪
                        g_SankyuInfo(w_numLoop).strTwinFlg = General.paGetDbFieldVal(w_TwinFlg_F, "")
                        '�o�Y�N����
                        g_SankyuInfo(w_numLoop).numBirthDate = Integer.Parse(General.paGetDbFieldVal(w_BirthDate_F, 0))
                        '�\��Y�O���From
                        g_SankyuInfo(w_numLoop).numPlanSanzenYamenFrom = Integer.Parse(General.paGetDbFieldVal(w_PlanSanzenYamenFrom_F, 0))
                        '�\��Y�O���To
                        g_SankyuInfo(w_numLoop).numPlanSanzenYamenTo = Integer.Parse(General.paGetDbFieldVal(w_PlanSanzenYamenTo_F, 99999999))
                        '�\��Y�O�x��From
                        g_SankyuInfo(w_numLoop).numPlanSanzenHolFrom = Integer.Parse(General.paGetDbFieldVal(w_PlanSanzenHolFrom_F, 0))
                        '�\��Y�O�x��To
                        g_SankyuInfo(w_numLoop).numPlanSanzenHolTo = Integer.Parse(General.paGetDbFieldVal(w_PlanSanzenHolTo_F, 99999999))
                        '�\��Y��x��From
                        g_SankyuInfo(w_numLoop).numPlanSangoHolFrom = Integer.Parse(General.paGetDbFieldVal(w_PlanSangoHolFrom_F, 0))
                        '�\��Y��x��To
                        g_SankyuInfo(w_numLoop).numPlanSangoHolTo = Integer.Parse(General.paGetDbFieldVal(w_PlanSangoHolTo_F, 99999999))
                        '�\��玙�x��From
                        g_SankyuInfo(w_numLoop).numPlanIkujiHolFrom = Integer.Parse(General.paGetDbFieldVal(w_PlanIkujiHolFrom_F, 0))
                        '�\��玙�x��To
                        g_SankyuInfo(w_numLoop).numPlanIkujiHolTo = Integer.Parse(General.paGetDbFieldVal(w_PlanIkujiHolTo_F, 99999999))
                        '�m��Y�O���From
                        g_SankyuInfo(w_numLoop).numFixedSanzenYamenFrom = Integer.Parse(General.paGetDbFieldVal(w_FixedSanzenYamenFrom_F, 0))
                        '�m��Y�O���To
                        g_SankyuInfo(w_numLoop).numFixedSanzenYamenTo = Integer.Parse(General.paGetDbFieldVal(w_FixedSanzenYamenTo_F, 99999999))
                        '�m��Y�O�x��From
                        g_SankyuInfo(w_numLoop).numFixedSanzenHolFrom = Integer.Parse(General.paGetDbFieldVal(w_FixedSanzenHolFrom_F, 0))
                        '�m��Y�O�x��To
                        g_SankyuInfo(w_numLoop).numFixedSanzenHolTo = Integer.Parse(General.paGetDbFieldVal(w_FixedSanzenHolTo_F, 99999999))
                        '�m��Y��x��From
                        g_SankyuInfo(w_numLoop).numFixedSangoHolFrom = Integer.Parse(General.paGetDbFieldVal(w_FixedSangoHolFrom_F, 0))
                        '�m��Y��x��To
                        g_SankyuInfo(w_numLoop).numFixedSangoHolTo = Integer.Parse(General.paGetDbFieldVal(w_FixedSangoHolTo_F, 99999999))
                        '�m��玙�x��From
                        g_SankyuInfo(w_numLoop).numFixedIkujiHolFrom = Integer.Parse(General.paGetDbFieldVal(w_FixedIkujiHolFrom_F, 0))
                        '�m��玙�x��To
                        g_SankyuInfo(w_numLoop).numFixedIkujiHolTo = Integer.Parse(General.paGetDbFieldVal(w_FixedIkujiHolTo_F, 99999999))
                        'UNIQUESEQNO
                        g_SankyuInfo(w_numLoop).strUniqueSeqNO = ""
                        '���F�σt���O
                        g_SankyuInfo(w_numLoop).strApproveFlg = ""
                        '����o�^����
                        g_SankyuInfo(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_SankyuInfo(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetSankyuInfo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' ��������擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetKyoukaiInfo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetKyoukaiInfo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Const W_strKyoukai As String = "S007" '����̗p�@�ėp�}�X�^ID

        Dim w_Rs As ADODB.Recordset
        Dim w_CD_F As ADODB.Field '�R�[�h
        Dim w_Name_F As ADODB.Field '����
        Dim w_Date_F As ADODB.Field '����N����
        Dim w_No_F As ADODB.Field '����ԍ�
        Dim w_FirstTime_F As ADODB.Field '����o�^����
        Dim w_UpdTime_F As ADODB.Field '�ŏI�X�V����
        Dim w_WithDrawDate_F As ADODB.Field '�މ�N����

        fncGetKyoukaiInfo = False
        Try
            ReDim g_KyoukaiInfo(0)

            w_strSql = ""
            w_strSql = w_strSql & " SELECT SO.SOCIETYCD "
            w_strSql = w_strSql & " ,      HM.NAME "
            w_strSql = w_strSql & " ,      SO.ENTERDATE "
            w_strSql = w_strSql & " ,      SO.WITHDRAWDATE " '�V���ɑމ�N�������擾
            w_strSql = w_strSql & " ,      SO.SOCIETYNO "
            w_strSql = w_strSql & " ,      SO.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      SO.LASTUPDTIMEDATE "

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals((General.gInstall_Enum.AccessType_PassThrough).ToString) Then 'ORACLE

                w_strSql = w_strSql & " FROM   NS_SOCIETYINFO_F SO "
                w_strSql = w_strSql & " ,      NS_HANYOU_M      HM "
                w_strSql = w_strSql & " WHERE  SO.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    SO.STAFFMNGID    = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    HM.HOSPITALCD  (+)  = SO.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM.MASTERID    (+)  = '" & W_strKyoukai & "' "
                w_strSql = w_strSql & " AND    HM.MASTERCD    (+)  = SO.SOCIETYCD "

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_SOCIETYINFO_F SO "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M HM "
                w_strSql = w_strSql & " ON     HM.HOSPITALCD    = SO.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM.MASTERID      = '" & W_strKyoukai & "' "
                w_strSql = w_strSql & " AND    HM.MASTERCD      = SO.SOCIETYCD "
                w_strSql = w_strSql & " WHERE  SO.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    SO.STAFFMNGID    = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND SO.ENTERDATE   <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND SO.ENTERDATE   >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " AND SO.ENTERDATE   <= " & m_numDateTo & " "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY SO.ENTERDATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY SO.ENTERDATE DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numKyoukaiKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numKyoukaiKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_KyoukaiInfo(m_numKyoukaiKensu)

                    w_CD_F = .Fields("SOCIETYCD") '�R�[�h
                    w_Name_F = .Fields("NAME") '����
                    w_Date_F = .Fields("ENTERDATE") '����N����
                    w_No_F = .Fields("SOCIETYNO") '����ԍ�
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE") '����o�^����
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE") '�ŏI�X�V����

                    w_WithDrawDate_F = .Fields("WITHDRAWDATE") '�މ�N����

                    For w_numLoop = 1 To m_numKyoukaiKensu

                        g_KyoukaiInfo(w_numLoop).strHospitalCD = m_strHospitalCD
                        g_KyoukaiInfo(w_numLoop).strStaffMngID = m_strStaffMngID

                        '�R�[�h

                        g_KyoukaiInfo(w_numLoop).strCD = General.paGetDbFieldVal(w_CD_F, "")
                        '����
                        g_KyoukaiInfo(w_numLoop).strName = General.paGetDbFieldVal(w_Name_F, "")
                        '����N����
                        g_KyoukaiInfo(w_numLoop).numGetDate = Integer.Parse(General.paGetDbFieldVal(w_Date_F, 0))
                        '�މ�N����
                        g_KyoukaiInfo(w_numLoop).numEndDate = Integer.Parse(General.paGetDbFieldVal(w_WithDrawDate_F, 99999999))
                        '����ԍ�
                        g_KyoukaiInfo(w_numLoop).strNo = General.paGetDbFieldVal(w_No_F, "")
                        '����o�^����
                        g_KyoukaiInfo(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_KyoukaiInfo(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))

                        g_KyoukaiInfo(w_numLoop).numDateFrom = 0
                        g_KyoukaiInfo(w_numLoop).numDateTo = 0

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetKyoukaiInfo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' �Ƒ������擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetKazokuInfo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetKazokuInfo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Const W_strKazoku As String = "S003" '�Ƒ����̗p�@�ėp�}�X�^ID

        Dim w_Rs As ADODB.Recordset
        Dim w_Name_F As ADODB.Field '����
        Dim w_SEQ_F As ADODB.Field 'SEQ
        Dim w_Date_F As ADODB.Field '���N����
        Dim w_TsudukiCD_F As ADODB.Field '�����R�[�h
        Dim w_TsudukiName_F As ADODB.Field '��������
        Dim w_Doukyo_F As ADODB.Field '�����敪
        Dim w_Fuyou_F As ADODB.Field '�}�{�敪
        Dim w_Seizon_F As ADODB.Field '�����敪
        Dim w_FirstTime_F As ADODB.Field '����o�^����
        Dim w_UpdTime_F As ADODB.Field '�ŏI�X�V����


        fncGetKazokuInfo = False
        Try
            ReDim g_KazokuInfo(0)

            w_strSql = ""
            w_strSql = w_strSql & " SELECT FY.FAMILYNAME "
            w_strSql = w_strSql & " ,      FY.RELATIONSHIPCD "
            w_strSql = w_strSql & " ,      HM.NAME "
            w_strSql = w_strSql & " ,      FY.SEQ "
            w_strSql = w_strSql & " ,      FY.BIRTHDAY "
            w_strSql = w_strSql & " ,      FY.COHABIKBN "
            w_strSql = w_strSql & " ,      FY.SUPPORTKBN "
            w_strSql = w_strSql & " ,      FY.LIVKBN "
            w_strSql = w_strSql & " ,      FY.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      FY.LASTUPDTIMEDATE "

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals((General.gInstall_Enum.AccessType_PassThrough).ToString) Then 'ORACLE

                w_strSql = w_strSql & " FROM   NS_FAMILYINFO_F FY "
                w_strSql = w_strSql & " ,      NS_HANYOU_M     HM "
                w_strSql = w_strSql & " WHERE  FY.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    FY.STAFFMNGID    = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    HM.HOSPITALCD  (+)  = FY.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM.MASTERID    (+)  = '" & W_strKazoku & "' "
                w_strSql = w_strSql & " AND    HM.MASTERCD    (+)  = FY.RELATIONSHIPCD "

            Else '����ȊO

                w_strSql = w_strSql & " FROM   NS_FAMILYINFO_F FY "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M HM "
                w_strSql = w_strSql & " ON     HM.HOSPITALCD    = FY.HOSPITALCD "
                w_strSql = w_strSql & " AND    HM.MASTERID      = '" & W_strKazoku & "' "
                w_strSql = w_strSql & " AND    HM.MASTERCD      = FY.RELATIONSHIPCD "
                w_strSql = w_strSql & " WHERE  FY.HOSPITALCD    = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    FY.STAFFMNGID    = '" & m_strStaffMngID & "' "

            End If

            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND FY.BIRTHDAY   <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND FY.BIRTHDAY   >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " AND FY.BIRTHDAY   <= " & m_numDateTo & " "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY FY.SEQ ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY FY.SEQ DESC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numKazokuKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numKazokuKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_KazokuInfo(m_numKazokuKensu)

                    w_Name_F = .Fields("FAMILYNAME") '����
                    w_Date_F = .Fields("BIRTHDAY") '���N����
                    w_TsudukiCD_F = .Fields("RELATIONSHIPCD") '�����R�[�h
                    w_TsudukiName_F = .Fields("NAME") '��������
                    w_SEQ_F = .Fields("SEQ") 'SEQ
                    w_Doukyo_F = .Fields("COHABIKBN") '�����敪
                    w_Fuyou_F = .Fields("SUPPORTKBN") '�}�{�敪
                    w_Seizon_F = .Fields("LIVKBN") '�����敪
                    w_FirstTime_F = .Fields("REGISTFIRSTTIMEDATE") '����o�^����
                    w_UpdTime_F = .Fields("LASTUPDTIMEDATE") '�ŏI�X�V����

                    For w_numLoop = 1 To m_numKazokuKensu

                        g_KazokuInfo(w_numLoop).strHospitalCD = m_strHospitalCD
                        g_KazokuInfo(w_numLoop).strStaffMngID = m_strStaffMngID

                        '����
                        g_KazokuInfo(w_numLoop).strName = General.paGetDbFieldVal(w_Name_F, "")
                        '���N����
                        g_KazokuInfo(w_numLoop).numDate = Integer.Parse(General.paGetDbFieldVal(w_Date_F, 0))
                        '�����R�[�h
                        g_KazokuInfo(w_numLoop).strTsudukiCD = General.paGetDbFieldVal(w_TsudukiCD_F, "")
                        '��������
                        g_KazokuInfo(w_numLoop).strTsudukiName = General.paGetDbFieldVal(w_TsudukiName_F, "")
                        'SEQ
                        g_KazokuInfo(w_numLoop).numSEQ = Integer.Parse(General.paGetDbFieldVal(w_SEQ_F, 0))
                        '�����敪
                        g_KazokuInfo(w_numLoop).strDoukyoKbn = General.paGetDbFieldVal(w_Doukyo_F, "")
                        '�}�{�敪
                        g_KazokuInfo(w_numLoop).strFuyouKbn = General.paGetDbFieldVal(w_Fuyou_F, "")
                        '�����敪
                        g_KazokuInfo(w_numLoop).strSeizonKbn = General.paGetDbFieldVal(w_Seizon_F, "")
                        '����o�^����
                        g_KazokuInfo(w_numLoop).lngFirstTime = Long.Parse(General.paGetDbFieldVal(w_FirstTime_F, 0))
                        '�ŏI�X�V����
                        g_KazokuInfo(w_numLoop).lngUpdTime = Long.Parse(General.paGetDbFieldVal(w_UpdTime_F, 0))

                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With

            w_Rs = Nothing

            fncGetKazokuInfo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' ���C�����擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[����AFALSE�F�G���[�Ȃ��j</returns>
    ''' <remarks></remarks>
    Public Function fncGetStudyInfo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetStudyInfo"

        Dim w_strSql As String
        Dim w_numLoop As Integer


        Dim W_strDivisionCD As String '���C�敪�p�@�ėp�}�X�^ID
        Dim W_strKindCD As String '���C��ޗp�@�ėp�}�X�^ID
        Dim W_strSponsorCD As String '���C��×p�@�ėp�}�X�^ID
        Dim W_strJoinCD As String '���C�Q���`�ԗp�@�ėp�}�X�^ID
        Dim W_strCourseCD As String '���C�R�[�X�p�@�ėp�}�X�^ID
        Dim W_strCostCD As String '��p�敪�p�@�ėp�}�X�^ID

        Dim w_Rs As ADODB.Recordset
        'Ͻ�����擾�����e����

        fncGetStudyInfo = False
        Try

            W_strCourseCD = General.paGetItemValue(General.G_StrMainKey4, General.G_StrSubKey11, "COURSECD", "", m_strHospitalCD) '�ėp�l�R�[�XCD�ݒ�
            W_strDivisionCD = General.paGetItemValue(General.G_StrMainKey4, General.G_StrSubKey11, "DIVISIONCD", "", m_strHospitalCD) '�ėp�l�敪CD�ݒ�
            W_strKindCD = General.paGetItemValue(General.G_StrMainKey4, General.G_StrSubKey11, "KINDCD", "", m_strHospitalCD) '�ėp�l���CD�ݒ�
            W_strSponsorCD = General.paGetItemValue(General.G_StrMainKey4, General.G_StrSubKey11, "SPONSORCD", "", m_strHospitalCD) '�ėp�l���CD�ݒ�
            W_strJoinCD = General.paGetItemValue(General.G_StrMainKey4, General.G_StrSubKey11, "JOINCD", "", m_strHospitalCD) '�ėp�l�Q���`��CD�ݒ�
            W_strCostCD = General.paGetItemValue(General.G_StrMainKey4, General.G_StrSubKey11, "COSTCD", "", m_strHospitalCD) '�ėp�l�̗p�敪CD�ݒ�
            
            ReDim g_StudyInfo(0)

            w_strSql = ""
            w_strSql = w_strSql & "SELECT KEN.NENDO, KEN.STUDYIDX, " & vbCr
            w_strSql = w_strSql & "       KEN.COURSECD,    COU.NAME AS COUNAME, " & vbCr
            w_strSql = w_strSql & "       KEN.KBNCD,       DIV.NAME AS DIVNAME, " & vbCr
            w_strSql = w_strSql & "       KEN.KINDCD,      KIN.NAME AS KINNAME, " & vbCr
            w_strSql = w_strSql & "       KEN.SPONSORCD,   SPO.NAME AS SPONAME, " & vbCr
            w_strSql = w_strSql & "       KEN.SANKAFORMCD, JOI.NAME AS JOINAME, " & vbCr
            w_strSql = w_strSql & "       MOU.COSTCD,      COS.NAME AS COSNAME  " & vbCr
            w_strSql = w_strSql & "     , COS2.NAME AS COSNAME2  " & vbCr
            w_strSql = w_strSql & "     , MOU.ATTENDLECSTATE, MOU.ATTENDLECREP, MOU.BIKOU " & vbCr
            w_strSql = w_strSql & "     , KEN.THEME, KEN.REPORTS " & vbCr
            w_strSql = w_strSql & "     , MOU.DELKBN " & vbCr
            w_strSql = w_strSql & "     , KEN.NENDOPLANFLG " & vbCr

            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals((General.gInstall_Enum.AccessType_PassThrough).ToString) Then 'ORACLE

                w_strSql = w_strSql & "FROM NS_STUDY_F KEN, NS_STUDYAPPLI_F MOU " & vbCr
                w_strSql = w_strSql & "   , NS_HANYOU_M COU, NS_HANYOU_M DIV " & vbCr
                w_strSql = w_strSql & "   , NS_HANYOU_M KIN, NS_HANYOU_M SPO " & vbCr
                w_strSql = w_strSql & "   , NS_HANYOU_M JOI, NS_HANYOU_M COS " & vbCr
                w_strSql = w_strSql & "   , NS_HANYOU_M COS2 " & vbCr
                w_strSql = w_strSql & "   , (SELECT HOSPITALCD, NENDO, STUDYIDX, MIN(DATEFROM) AS DATEFROM " & vbCr
                w_strSql = w_strSql & "     FROM NS_STUDYDATE_F " & vbCr
                w_strSql = w_strSql & "     WHERE HOSPITALCD = '" & m_strHospitalCD & "' " & vbCr
                '�N�x�̎w�肪����ꍇ
                If m_numNendo <> 0 Then
                    w_strSql = w_strSql & " AND NENDO = " & m_numNendo & " " & vbCr
                End If
                w_strSql = w_strSql & "     GROUP BY HOSPITALCD, NENDO, STUDYIDX) SD1 " & vbCr
                w_strSql = w_strSql & "   , (SELECT HOSPITALCD, NENDO, STUDYIDX " & vbCr
                w_strSql = w_strSql & "     FROM NS_STUDYDATE_F " & vbCr
                w_strSql = w_strSql & "     WHERE HOSPITALCD = '" & m_strHospitalCD & "' " & vbCr
                If m_numNendo <> 0 Then
                    w_strSql = w_strSql & " AND NENDO = " & m_numNendo & " " & vbCr
                End If
                '�P����̏ꍇ
                If m_numDateFlg = 0 Then
                    w_strSql = w_strSql & " AND DATEFROM <= " & m_numDateFrom & " " & vbCr
                    '���Ԏw��̏ꍇ
                ElseIf m_numDateFlg = 1 Then
                    '�P����̌��C�̏ꍇ�ɏI������ 0 �ɂȂ�̂ŁA���̑Ή���ǉ�
                    w_strSql = w_strSql & " AND (( DATEFROM <= " & m_numDateTo & " " & vbCr
                    w_strSql = w_strSql & " AND DATETO >= " & m_numDateFrom & ") " & vbCr
                    w_strSql = w_strSql & " OR  ( DATEFROM <= " & m_numDateTo & " " & vbCr
                    w_strSql = w_strSql & " AND DATEFROM >= " & m_numDateFrom & ")) " & vbCr
                End If
                w_strSql = w_strSql & "     GROUP BY HOSPITALCD, NENDO, STUDYIDX) SD2 " & vbCr
                w_strSql = w_strSql & "WHERE MOU.HOSPITALCD   = '" & m_strHospitalCD & "' " & vbCr
                w_strSql = w_strSql & " AND  KEN.STUDYIDX     = MOU.STUDYIDX " & vbCr
                w_strSql = w_strSql & " AND  KEN.NENDO        = MOU.NENDO " & vbCr
                w_strSql = w_strSql & " AND  KEN.HOSPITALCD   = MOU.HOSPITALCD " & vbCr
                w_strSql = w_strSql & " AND  MOU.STAFFMNGID   = '" & m_strStaffMngID & "' " & vbCr
                '�N�x�̎w�肪����ꍇ
                If m_numNendo <> 0 Then
                    w_strSql = w_strSql & " AND  MOU.NENDO    = " & m_numNendo & vbCr
                End If
                '�폜�󋵂��w��
                If m_intDelKbn <> 2 Then
                    w_strSql = w_strSql & " AND  MOU.DELKBN   = " & m_intDelKbn & vbCr
                End If
                w_strSql = w_strSql & " AND  SD1.HOSPITALCD   = MOU.HOSPITALCD " & vbCr
                w_strSql = w_strSql & " AND  SD1.NENDO        = MOU.NENDO " & vbCr
                w_strSql = w_strSql & " AND  SD1.STUDYIDX     = MOU.STUDYIDX " & vbCr
                w_strSql = w_strSql & " AND  SD2.HOSPITALCD   = MOU.HOSPITALCD " & vbCr
                w_strSql = w_strSql & " AND  SD2.NENDO        = MOU.NENDO " & vbCr
                w_strSql = w_strSql & " AND  SD2.STUDYIDX     = MOU.STUDYIDX " & vbCr

                '�ėp�l�R�[�XCD�ݒ�
                w_strSql = w_strSql & " AND    COU.HOSPITALCD  (+)  = KEN.HOSPITALCD "
                w_strSql = w_strSql & " AND    COU.MASTERID    (+)  = '" & W_strCourseCD & "' "
                w_strSql = w_strSql & " AND    COU.MASTERCD    (+)  = KEN.COURSECD "
                '�ėp�l�敪CD�ݒ�
                w_strSql = w_strSql & " AND    DIV.HOSPITALCD  (+)  = KEN.HOSPITALCD "
                w_strSql = w_strSql & " AND    DIV.MASTERID    (+)  = '" & W_strDivisionCD & "' "
                w_strSql = w_strSql & " AND    DIV.MASTERCD    (+)  = KEN.KBNCD "
                '�ėp�l���CD�ݒ�
                w_strSql = w_strSql & " AND    KIN.HOSPITALCD  (+)  = KEN.HOSPITALCD "
                w_strSql = w_strSql & " AND    KIN.MASTERID    (+)  = '" & W_strKindCD & "' "
                w_strSql = w_strSql & " AND    KIN.MASTERCD    (+)  = KEN.KINDCD "
                '�ėp�l���CD�ݒ�
                w_strSql = w_strSql & " AND    SPO.HOSPITALCD  (+)  = KEN.HOSPITALCD "
                w_strSql = w_strSql & " AND    SPO.MASTERID    (+)  = '" & W_strSponsorCD & "' "
                w_strSql = w_strSql & " AND    SPO.MASTERCD    (+)  = KEN.SPONSORCD "
                '�ėp�l�Q���`��CD�ݒ�
                w_strSql = w_strSql & " AND    JOI.HOSPITALCD  (+)  = KEN.HOSPITALCD "
                w_strSql = w_strSql & " AND    JOI.MASTERID    (+)  = '" & W_strJoinCD & "' "
                w_strSql = w_strSql & " AND    JOI.MASTERCD    (+)  = KEN.SANKAFORMCD "
                '�ėp�l��p�敪CD�ݒ�
                w_strSql = w_strSql & " AND    COS.HOSPITALCD  (+)  = KEN.HOSPITALCD "
                w_strSql = w_strSql & " AND    COS.MASTERID    (+)  = '" & W_strCostCD & "' "
                w_strSql = w_strSql & " AND    COS.MASTERCD    (+)  = KEN.COSTCD "
                '�ėp�l��p�敪CD�ݒ�(���C�\��F)
                w_strSql = w_strSql & " AND    COS2.HOSPITALCD  (+)  = MOU.HOSPITALCD "
                w_strSql = w_strSql & " AND    COS2.MASTERID    (+)  = '" & W_strCostCD & "' "
                w_strSql = w_strSql & " AND    COS2.MASTERCD    (+)  = MOU.COSTCD "

            Else '����ȊO

                w_strSql = w_strSql & "FROM NS_STUDY_F KEN " & vbCr

                '�ėp�l�R�[�XCD�ݒ�
                w_strSql = w_strSql & "     LEFT OUTER JOIN NS_HANYOU_M COU " & vbCr
                w_strSql = w_strSql & "      ON     COU.HOSPITALCD   = KEN.HOSPITALCD " & vbCr
                w_strSql = w_strSql & "     AND    COU.MASTERID      = '" & W_strCourseCD & "' " & vbCr
                w_strSql = w_strSql & "     AND    COU.MASTERCD      = KEN.COURSECD " & vbCr

                '�ėp�l�敪CD�ݒ�
                w_strSql = w_strSql & "     LEFT OUTER JOIN NS_HANYOU_M DIV " & vbCr
                w_strSql = w_strSql & "      ON     DIV.HOSPITALCD   = KEN.HOSPITALCD " & vbCr
                w_strSql = w_strSql & "     AND    DIV.MASTERID      = '" & W_strDivisionCD & "' " & vbCr
                w_strSql = w_strSql & "     AND    DIV.MASTERCD      = KEN.KBNCD " & vbCr

                '�ėp�l���CD�ݒ�
                w_strSql = w_strSql & "     LEFT OUTER JOIN NS_HANYOU_M KIN " & vbCr
                w_strSql = w_strSql & "      ON     KIN.HOSPITALCD   = KEN.HOSPITALCD " & vbCr
                w_strSql = w_strSql & "     AND    KIN.MASTERID      = '" & W_strKindCD & "' " & vbCr
                w_strSql = w_strSql & "     AND    KIN.MASTERCD      = KEN.KINDCD " & vbCr

                '�ėp�l���CD�ݒ�
                w_strSql = w_strSql & "     LEFT OUTER JOIN NS_HANYOU_M SPO " & vbCr
                w_strSql = w_strSql & "      ON     SPO.HOSPITALCD   = KEN.HOSPITALCD " & vbCr
                w_strSql = w_strSql & "     AND    SPO.MASTERID      = '" & W_strSponsorCD & "' " & vbCr
                w_strSql = w_strSql & "     AND    SPO.MASTERCD      = KEN.SPONSORCD " & vbCr

                '�ėp�l�Q���`��CD�ݒ�
                w_strSql = w_strSql & "     LEFT OUTER JOIN NS_HANYOU_M JOI" & vbCr
                w_strSql = w_strSql & "      ON     JOI.HOSPITALCD   = KEN.HOSPITALCD " & vbCr
                w_strSql = w_strSql & "     AND    JOI.MASTERID      = '" & W_strJoinCD & "' " & vbCr
                w_strSql = w_strSql & "     AND    JOI.MASTERCD      = KEN.SANKAFORMCD " & vbCr

                '�ėp�l��p�敪CD�ݒ�
                w_strSql = w_strSql & "     LEFT OUTER JOIN NS_HANYOU_M COS " & vbCr
                w_strSql = w_strSql & "      ON     COS.HOSPITALCD   = KEN.HOSPITALCD " & vbCr
                w_strSql = w_strSql & "     AND    COS.MASTERID      = '" & W_strCostCD & "' " & vbCr
                w_strSql = w_strSql & "     AND    COS.MASTERCD      = KEN.COSTCD " & vbCr

                w_strSql = w_strSql & "   , NS_STUDYAPPLI_F MOU " & vbCr
                w_strSql = w_strSql & "     INNER JOIN (SELECT HOSPITALCD, NENDO, STUDYIDX, MIN(DATEFROM) AS DATEFROM " & vbCr
                w_strSql = w_strSql & "     FROM NS_STUDYDATE_F " & vbCr
                w_strSql = w_strSql & "     WHERE HOSPITALCD = '" & m_strHospitalCD & "' " & vbCr
                '�N�x�̎w�肪����ꍇ
                If m_numNendo <> 0 Then
                    w_strSql = w_strSql & " AND NENDO = " & m_numNendo & " " & vbCr
                End If
                w_strSql = w_strSql & "     GROUP BY HOSPITALCD, NENDO, STUDYIDX) SD1 " & vbCr
                w_strSql = w_strSql & "     ON  SD1.HOSPITALCD   = MOU.HOSPITALCD " & vbCr
                w_strSql = w_strSql & "     AND SD1.NENDO        = MOU.NENDO " & vbCr
                w_strSql = w_strSql & "     AND SD1.STUDYIDX     = MOU.STUDYIDX " & vbCr

                w_strSql = w_strSql & "     INNER JOIN (SELECT HOSPITALCD, NENDO, STUDYIDX " & vbCr
                w_strSql = w_strSql & "     FROM NS_STUDYDATE_F " & vbCr
                w_strSql = w_strSql & "     WHERE HOSPITALCD = '" & m_strHospitalCD & "' " & vbCr
                If m_numNendo <> 0 Then
                    w_strSql = w_strSql & " AND NENDO = " & m_numNendo & " " & vbCr
                End If
                '�P����̏ꍇ
                If m_numDateFlg = 0 Then
                    w_strSql = w_strSql & " AND DATEFROM <= " & m_numDateFrom & " " & vbCr
                    '���Ԏw��̏ꍇ
                ElseIf m_numDateFlg = 1 Then
                    '�P����̌��C�̏ꍇ�ɏI������ 0 �ɂȂ�̂ŁA���̑Ή���ǉ�
                    w_strSql = w_strSql & " AND (( DATEFROM <= " & m_numDateTo & " " & vbCr
                    w_strSql = w_strSql & " AND DATETO >= " & m_numDateFrom & ") " & vbCr
                    w_strSql = w_strSql & " OR  ( DATEFROM <= " & m_numDateTo & " " & vbCr
                    w_strSql = w_strSql & " AND DATEFROM >= " & m_numDateFrom & ")) " & vbCr
                End If
                w_strSql = w_strSql & "     GROUP BY HOSPITALCD, NENDO, STUDYIDX) SD2 " & vbCr
                w_strSql = w_strSql & "     ON  SD2.HOSPITALCD   = MOU.HOSPITALCD " & vbCr
                w_strSql = w_strSql & "     AND SD2.NENDO        = MOU.NENDO " & vbCr
                w_strSql = w_strSql & "     AND SD2.STUDYIDX     = MOU.STUDYIDX " & vbCr
                '�ėp�l��p�敪CD�ݒ�(���C�\��F)
                w_strSql = w_strSql & "     LEFT OUTER JOIN NS_HANYOU_M COS2 " & vbCr
                w_strSql = w_strSql & "      ON    COS2.HOSPITALCD   = MOU.HOSPITALCD " & vbCr
                w_strSql = w_strSql & "     AND    COS2.MASTERID     = '" & W_strCostCD & "' " & vbCr
                w_strSql = w_strSql & "     AND    COS2.MASTERCD     = MOU.COSTCD " & vbCr
                w_strSql = w_strSql & "WHERE KEN.HOSPITALCD   = '" & m_strHospitalCD & "' " & vbCr
                w_strSql = w_strSql & " AND  KEN.STUDYIDX     = MOU.STUDYIDX " & vbCr
                w_strSql = w_strSql & " AND  KEN.NENDO        = MOU.NENDO " & vbCr
                w_strSql = w_strSql & " AND  MOU.HOSPITALCD   = KEN.HOSPITALCD " & vbCr
                w_strSql = w_strSql & " AND  MOU.STAFFMNGID   = '" & m_strStaffMngID & "' " & vbCr
                '�N�x�̎w�肪����ꍇ
                If m_numNendo <> 0 Then
                    w_strSql = w_strSql & " AND MOU.NENDO = " & m_numNendo & " " & vbCr
                End If
                '�폜�󋵂��w��
                If m_intDelKbn <> 2 Then
                    w_strSql = w_strSql & " AND  MOU.DELKBN   = " & m_intDelKbn & vbCr
                End If
            End If
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & "ORDER BY MOU.NENDO, SD1.DATEFROM, MOU.STUDYIDX "
            Else
                w_strSql = w_strSql & "ORDER BY MOU.NENDO DESC, SD1.DATEFROM DESC, MOU.STUDYIDX DESC "
            End If
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numStudyKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numStudyKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_StudyInfo(m_numStudyKensu)

                    For w_numLoop = 1 To m_numStudyKensu
                        g_StudyInfo(w_numLoop).numYYYY = Integer.Parse(General.paGetDbFieldVal(.Fields("NENDO"), 0))
                        g_StudyInfo(w_numLoop).strSEQ = General.paGetDbFieldVal(.Fields("STUDYIDX"), "")
                        g_StudyInfo(w_numLoop).strCourseCD = General.paGetDbFieldVal(.Fields("COURSECD"), "")
                        g_StudyInfo(w_numLoop).strCorseName = General.paGetDbFieldVal(.Fields("COUNAME"), "")
                        g_StudyInfo(w_numLoop).strKbnCD = General.paGetDbFieldVal(.Fields("KBNCD"), "")
                        g_StudyInfo(w_numLoop).strKbnName = General.paGetDbFieldVal(.Fields("DIVNAME"), "")
                        g_StudyInfo(w_numLoop).strSyuruiCD = General.paGetDbFieldVal(.Fields("KINDCD"), "")
                        g_StudyInfo(w_numLoop).strSyuruiName = General.paGetDbFieldVal(.Fields("KINNAME"), "")
                        g_StudyInfo(w_numLoop).strSyusaiCD = General.paGetDbFieldVal(.Fields("SPONSORCD"), "")
                        g_StudyInfo(w_numLoop).strSyusaiName = General.paGetDbFieldVal(.Fields("SPONAME"), "")
                        g_StudyInfo(w_numLoop).strSankaCD = General.paGetDbFieldVal(.Fields("SANKAFORMCD"), "")
                        g_StudyInfo(w_numLoop).strSankaName = General.paGetDbFieldVal(.Fields("JOINAME"), "")
                        g_StudyInfo(w_numLoop).strCostCD = General.paGetDbFieldVal(.Fields("COSTCD"), "")
                        g_StudyInfo(w_numLoop).strCostName = General.paGetDbFieldVal(.Fields("COSNAME"), "")
                        g_StudyInfo(w_numLoop).strApplyStatus = General.paGetDbFieldVal(.Fields("ATTENDLECSTATE"), "")
                        g_StudyInfo(w_numLoop).strApplyRepo = General.paGetDbFieldVal(.Fields("ATTENDLECREP"), "")
                        g_StudyInfo(w_numLoop).strThema = General.paGetDbFieldVal(.Fields("THEME"), "")
                        g_StudyInfo(w_numLoop).strBiko = General.paGetDbFieldVal(.Fields("BIKOU"), "")
                        g_StudyInfo(w_numLoop).strDeleteStatus = Integer.Parse(General.paGetDbFieldVal(.Fields("DELKBN"), 0))
                        g_StudyInfo(w_numLoop).strPlaningFLG = Integer.Parse(General.paGetDbFieldVal(.Fields("NENDOPLANFLG"), 9))
                        g_StudyInfo(w_numLoop).strCostCD2 = General.paGetDbFieldVal(.Fields("COSTCD"), "")
                        g_StudyInfo(w_numLoop).strCostName2 = General.paGetDbFieldVal(.Fields("COSNAME2"), "")
                        .MoveNext()
                    Next w_numLoop
                End If
            End With

            w_Rs.Close()

            w_Rs = Nothing

            '���C�����̎擾
            Call fncGetStudyInfoSub()

            fncGetStudyInfo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ���C���t���̓ǂݍ���
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub fncGetStudyInfoSub()

        Const W_SUBNAME As String = "Nssk001 fncGetStudyInfoSub"

        Dim w_strSql As String
        Dim w_objRs As ADODB.Recordset
        Dim w_intYMDIndex As Short
        Dim w_intIndex As Short
        Dim w_strDateList As String
        Try
            For w_intYMDIndex = 1 To UBound(g_StudyInfo)
                '�\���p����������̏�����
                w_strDateList = ""

                ReDim g_StudyInfo(w_intYMDIndex).objDateList(0)

                w_strSql = ""
                w_strSql = w_strSql & "SELECT DATEAPPOFLG, DATEFROM, DATETO "
                w_strSql = w_strSql & " FROM NS_STUDYDATE_F "
                w_strSql = w_strSql & "WHERE STUDYIDX = " & CStr(g_StudyInfo(w_intYMDIndex).strSEQ) & " "
                w_strSql = w_strSql & "AND NENDO = " & CStr(g_StudyInfo(w_intYMDIndex).numYYYY) & " "
                w_strSql = w_strSql & "AND HOSPITALCD = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & "ORDER BY DATEIDX "

                w_objRs = General.paDBRecordSetOpen(w_strSql)

                If w_objRs.EOF = False Then
                    With w_objRs
                        .MoveLast()
                        m_numStudyDateKensu = .RecordCount

                        ReDim g_StudyInfo(w_intYMDIndex).objDateList(m_numStudyDateKensu)

                        .MoveFirst()

                        For w_intIndex = 1 To m_numStudyDateKensu
                            '�����f�[�^���m��
                            g_StudyInfo(w_intYMDIndex).objDateList(w_intIndex).strDateType = .Fields(0).Value & ""

                            g_StudyInfo(w_intYMDIndex).objDateList(w_intIndex).numFromDate = Integer.Parse(General.paGetDbFieldVal(.Fields(1), 0))

                            g_StudyInfo(w_intYMDIndex).objDateList(w_intIndex).numToDate = Integer.Parse(General.paGetDbFieldVal(.Fields(2), 0))

                            '�\���p������̕ҏW
                            Select Case .Fields(0).Value & ""
                                Case "0" '����
                                    If String.IsNullOrEmpty(w_strDateList) = False Then
                                        w_strDateList = w_strDateList & ","
                                    End If

                                    w_strDateList = w_strDateList & Format(Integer.Parse(General.paGetDbFieldVal(.Fields(1), 0)), "0000/00/00")
                                    w_strDateList = w_strDateList & "�`"

                                    w_strDateList = w_strDateList & Format(Integer.Parse(General.paGetDbFieldVal(.Fields(2), 0)), "0000/00/00")

                                Case "1" '�P���
                                    If String.IsNullOrEmpty(w_strDateList) = False Then
                                        w_strDateList = w_strDateList & ","
                                    End If

                                    w_strDateList = w_strDateList & Format(Integer.Parse(General.paGetDbFieldVal(.Fields(1), 0)), "0000/00/00")
                                    w_strDateList = w_strDateList & "�`" & w_strDateList '2012/10/25 fujisawa add 

                            End Select



                            .MoveNext()
                        Next w_intIndex

                        '2012/10/25 fujisawa add st ------------
                        '2���ȏ゠��ƈ����*��ǉ����ĕ\��
                        If w_intYMDIndex = 2 Then
                            w_strDateList = w_strDateList & " *"
                        End If
                        '2012/10/25 fujisawa add end ------------
                    End With
                Else
                    m_numStudyDateKensu = 0
                End If

                w_objRs.Close()

                w_objRs = Nothing

                '�\���p������̊m��
                g_StudyInfo(w_intYMDIndex).strDate = w_strDateList
            Next w_intYMDIndex
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End Try
    End Sub
    ''' <summary>
    ''' �Ɛя��擾
    ''' </summary>
    ''' <returns>�iTRUE�F�G���[�Ȃ��AFALSE�F�G���[����j</returns>
    ''' <remarks></remarks>
    Public Function fncGetGyosekiInfo() As Boolean

        Dim w_strPreErrorProc As String
        w_strPreErrorProc = General.g_ErrorProc
        General.g_ErrorProc = "BasNSC0060C fncGetGyosekiInfo"

        Dim w_strSql As String
        Dim w_numLoop As Integer

        Dim w_Rs As ADODB.Recordset

        Dim w_HOSPITALCD_F As ADODB.Field
        Dim w_StaffMngID_F As ADODB.Field
        Dim w_GyosekiCd_F As ADODB.Field
        Dim w_GyosekiName_F As ADODB.Field
        Dim w_SEQ_F As ADODB.Field
        Dim w_FromDate_F As ADODB.Field
        Dim w_ToDate_F As ADODB.Field
        Dim w_Subject_F As ADODB.Field
        Dim w_GyosekiPlaceCd_F As ADODB.Field
        Dim w_GyosekiPlaceName_F As ADODB.Field
        Dim w_GyosekiBikou_F As ADODB.Field
        Dim w_RegistFirstTimeDate_F As ADODB.Field
        Dim w_LastUpdTimeDate_F As ADODB.Field
        Const w_MSTID_B As String = "E002"
        Const w_MSTID_BP As String = "E008"

        fncGetGyosekiInfo = False
        Try
            ReDim g_Gyoseki(0)

            w_strSql = ""

            w_strSql = w_strSql & " SELECT "
            w_strSql = w_strSql & "        GI.HOSPITALCD "
            w_strSql = w_strSql & " ,      GI.STAFFMNGID "
            w_strSql = w_strSql & " ,      GI.GYOSEKICD "
            w_strSql = w_strSql & " ,      H1.NAME      AS GYOSEKINAME "
            w_strSql = w_strSql & " ,      GI.SEQ "
            w_strSql = w_strSql & " ,      GI.FROMDATE "
            w_strSql = w_strSql & " ,      GI.TODATE "
            w_strSql = w_strSql & " ,      GI.SUBJECT "
            w_strSql = w_strSql & " ,      GI.GYOSEKIPLACECD "
            w_strSql = w_strSql & " ,      H2.NAME      AS GYOSEKIPLACENAME "
            w_strSql = w_strSql & " ,      GI.GYOSEKIBIKOU "
            w_strSql = w_strSql & " ,      GI.REGISTFIRSTTIMEDATE "
            w_strSql = w_strSql & " ,      GI.LASTUPDTIMEDATE "
            w_strSql = w_strSql & " ,      GI.REGISTRANTID "
            w_strSql = w_strSql & " FROM   NS_GYOSEKIINFO_F GI "


            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals((General.gInstall_Enum.AccessType_PassThrough).ToString) Then 'ORACLE


                w_strSql = w_strSql & " ,      NS_HANYOU_M      H1 "
                w_strSql = w_strSql & " ,      NS_HANYOU_M      H2 "
                w_strSql = w_strSql & " WHERE  GI.HOSPITALCD     = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    GI.STAFFMNGID     = '" & m_strStaffMngID & "' "
                w_strSql = w_strSql & " AND    H1.HOSPITALCD (+) = GI.HOSPITALCD "
                w_strSql = w_strSql & " AND    H1.MASTERID (+)   = '" & w_MSTID_B & "' "
                w_strSql = w_strSql & " AND    H1.MASTERCD (+)   = GI.GYOSEKICD "
                w_strSql = w_strSql & " AND    H2.HOSPITALCD (+) = GI.HOSPITALCD "
                w_strSql = w_strSql & " AND    H2.MASTERID (+)   = '" & w_MSTID_BP & "' "
                w_strSql = w_strSql & " AND    H2.MASTERCD (+)   = GI.GYOSEKIPLACECD "

            Else '����ȊO

                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M H1 "
                w_strSql = w_strSql & " ON   H1.HOSPITALCD = GI.HOSPITALCD "
                w_strSql = w_strSql & " AND  H1.MASTERID   = '" & w_MSTID_B & "' "
                w_strSql = w_strSql & " AND  H1.MASTERCD   = GI.GYOSEKICD "
                w_strSql = w_strSql & " LEFT OUTER JOIN NS_HANYOU_M H2 "
                w_strSql = w_strSql & " ON   H2.HOSPITALCD = GI.HOSPITALCD "
                w_strSql = w_strSql & " AND  H2.MASTERID   = '" & w_MSTID_BP & "' "
                w_strSql = w_strSql & " AND  H2.MASTERCD   = GI.GYOSEKIPLACECD "
                w_strSql = w_strSql & " WHERE  GI.HOSPITALCD     = '" & m_strHospitalCD & "' "
                w_strSql = w_strSql & " AND    GI.STAFFMNGID     = '" & m_strStaffMngID & "' "

            End If
            '�P����̏ꍇ
            If m_numDateFlg = 0 Then
                w_strSql = w_strSql & " AND GI.FROMDATE      <= " & m_numDateFrom & " "
                '���Ԏw��̏ꍇ
            ElseIf m_numDateFlg = 1 Then
                w_strSql = w_strSql & " AND   GI.FROMDATE    <= " & m_numDateTo & " "
                w_strSql = w_strSql & " AND ( GI.TODATE >= " & m_numDateFrom & " "
                w_strSql = w_strSql & " OR    GI.TODATE  = 0 "
                w_strSql = w_strSql & " OR    GI.TODATE IS NULL ) "
            End If
            '����
            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY GI.FROMDATE ASC "
                '�~��
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY GI.FROMDATE DESC "
            End If


            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numGyosekiKensu = 0
                    .Close()

                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numGyosekiKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_Gyoseki(m_numGyosekiKensu)

                    w_HOSPITALCD_F = .Fields("HOSPITALCD")
                    w_StaffMngID_F = .Fields("STAFFMNGID")
                    w_GyosekiCd_F = .Fields("GYOSEKICD")
                    w_GyosekiName_F = .Fields("GYOSEKINAME")
                    w_SEQ_F = .Fields("SEQ")
                    w_FromDate_F = .Fields("FROMDATE")
                    w_ToDate_F = .Fields("TODATE")
                    w_Subject_F = .Fields("SUBJECT")
                    w_GyosekiPlaceCd_F = .Fields("GYOSEKIPLACECD")
                    w_GyosekiPlaceName_F = .Fields("GYOSEKIPLACENAME")
                    w_GyosekiBikou_F = .Fields("GYOSEKIBIKOU")
                    w_RegistFirstTimeDate_F = .Fields("REGISTFIRSTTIMEDATE")
                    w_LastUpdTimeDate_F = .Fields("LASTUPDTIMEDATE")


                    For w_numLoop = 1 To m_numGyosekiKensu

                        '�{�݃R�[�h
                        g_Gyoseki(w_numLoop).strHospitalCD = General.paGetDbFieldVal(w_HOSPITALCD_F, "")
                        '�E���Ǘ��ԍ�
                        g_Gyoseki(w_numLoop).strStaffMngID = General.paGetDbFieldVal(w_StaffMngID_F, "")
                        '�ƐуR�[�h
                        g_Gyoseki(w_numLoop).strGyosekiCd = General.paGetDbFieldVal(w_GyosekiCd_F, "")
                        '�Ɛі���
                        g_Gyoseki(w_numLoop).strGyosekiName = General.paGetDbFieldVal(w_GyosekiName_F, "")
                        'SEQ
                        g_Gyoseki(w_numLoop).numSEQ = Integer.Parse(General.paGetDbFieldVal(w_SEQ_F, 0))
                        '�J�n�N����
                        g_Gyoseki(w_numLoop).numFromDate = Integer.Parse(General.paGetDbFieldVal(w_FromDate_F, 0))
                        '�I���N����
                        g_Gyoseki(w_numLoop).numToDate = Integer.Parse(General.paGetDbFieldVal(w_ToDate_F, 99999999))
                        '����
                        g_Gyoseki(w_numLoop).strSubject = General.paGetDbFieldVal(w_Subject_F, "")
                        '�Ɛє��\�ꏊ�R�[�h
                        g_Gyoseki(w_numLoop).strGyosekiPlaceCd = General.paGetDbFieldVal(w_GyosekiPlaceCd_F, "")
                        '�Ɛє��\�ꏊ����
                        g_Gyoseki(w_numLoop).strGyosekiPlaceName = General.paGetDbFieldVal(w_GyosekiPlaceName_F, "")
                        '�Ɛє��l
                        g_Gyoseki(w_numLoop).strGyosekiBikou = General.paGetDbFieldVal(w_GyosekiBikou_F, "")
                        '����o�^����
                        g_Gyoseki(w_numLoop).dblRegistFirstTimeDate = Long.Parse(General.paGetDbFieldVal(w_RegistFirstTimeDate_F, 0))
                        '�ŏI�X�V����
                        g_Gyoseki(w_numLoop).dblLastUpdTimeDate = Long.Parse(General.paGetDbFieldVal(w_LastUpdTimeDate_F, 0))


                        .MoveNext()
                    Next w_numLoop
                End If
                .Close()
            End With
            w_Rs = Nothing

            fncGetGyosekiInfo = True

            General.g_ErrorProc = w_strPreErrorProc

        Catch ex As Exception

            Throw
        End Try
    End Function
    '2012/02/13 Sasaki add start-----------------------------------------------------------------------------------------------------------------------------------------
    '���C��u�������擾
    Public Function fncGetStudyAttend() As Boolean

        Dim w_intCount As Integer
        Dim w_intIndex As Integer
        Dim w_strSql As String
        Dim w_Rs As ADODB.Recordset
        Dim w_Nendo_F As ADODB.Field                            '�N�x                       �u�m�r���C�e
        Dim w_StudyIndex_F As ADODB.Field                       '���C�h�m�c�d�w             �u�m�r���C�e
        Dim w_OutInFlg_F As ADODB.Field                         '�@���O�t���O               �u�m�r���C�e
        Dim w_StudyCD_F As ADODB.Field                          '���C�R�[�h                 �u�m�r���C�e
        Dim w_StudyName_F As ADODB.Field                        '���C����                   �u�m�r���C�e
        Dim w_StudySecName_F As ADODB.Field                     '���C����                   �u�m�r���C�e
        Dim w_StudyKana_F As ADODB.Field                        '���C�t���K�i               �u�m�r���C�e
        Dim w_KindCD_F As ADODB.Field                           '��ރR�[�h                 �u�m�r���C�e
        Dim w_KindName_F As ADODB.Field                         '��ޖ���                   �u�m�r���C�e
        Dim w_SponsorCD_F As ADODB.Field                        '��ÃR�[�h                 �u�m�r���C�e
        Dim w_SponsorName_F As ADODB.Field                      '��Ö���                   �u�m�r���C�e
        Dim w_Theme_F As ADODB.Field                            '�e�[�}                     �u�m�r���C�e
        Dim w_Lecturer_F As ADODB.Field                         '�u�t                       �u�m�r���C�e
        Dim w_Hall_F As ADODB.Field                             '���E�ꏊ                 �u�m�r���C�e
        Dim w_SankaCond_F As ADODB.Field                        '�Q������                   �u�m�r���C�e
        Dim w_SankaNinzu_F As ADODB.Field                       '�Q���l��                   �u�m�r���C�e
        Dim w_Reports_F As ADODB.Field                          '�A������                   �u�m�r���C�e
        Dim w_Bikou_F As ADODB.Field                            '���l                       �u�m�r���C�e
        Dim w_Url_F As ADODB.Field                              '�t�q�k                     �u�m�r���C�e
        Dim w_NecessaryValuationLevelCD_F As ADODB.Field        '�K�{�]�����x���R�[�h       �u�m�r���C�e
        Dim w_NecessaryValuationLevelName_F As ADODB.Field      '�K�{�]�����x������         �u�m�r���C�e
        Dim w_NecessaryValuationLevelSecName_F As ADODB.Field   '�K�{�]�����x������         �u�m�r���C�e
        Dim w_NecessaryValuationLevelMark_F As ADODB.Field      '�K�{�]�����x���L��         �u�m�r���C�e
        Dim w_AcceptFromDate_F As ADODB.Field                   '��t�J�n�N����             �u�m�r���C�e
        Dim w_AcceptToDate_F As ADODB.Field                     '��t�I���N����             �u�m�r���C�e
        Dim w_Acceptapstate_F As ADODB.Field                    '��t�\���݃t���O           �u�m�r���C�e
        Dim w_NendoPlanKbn_F As ADODB.Field                     '�N�Ԍv��敪               �u�m�r���C�e
        Dim w_KinmuDeptCD_F As ADODB.Field                      '�Ζ������R�[�h             �u�m�r���C�e
        Dim w_KinmuDeptName_F As ADODB.Field                    '�Ζ���������               �u�m�r���C�e
        Dim w_AllDaysNecessaryFlg_F As ADODB.Field              '�S�����K�{�t���O           �u�m�r���C�e
        Dim w_IndependentFlg_F As ADODB.Field                   '���匤�C�t���O             �u�m�r���C�e
        Dim w_DateIdx_F As ADODB.Field                          '���t�C���f�b�N�X           �m�r���C���t�e
        Dim w_DateAppoFlg_F As ADODB.Field                      '���t�w��t���O             �m�r���C���t�e
        Dim w_DateFrom_F As ADODB.Field                         '���t�J�n�N����             �m�r���C���t�e
        Dim w_DateTo_F As ADODB.Field                           '���t�I���N����             �m�r���C���t�e
        Dim w_JapanAreaCD_F As ADODB.Field                      '�s���{���R�[�h             �m�r���C���t�e
        Dim w_JapanAreaName_F As ADODB.Field                    '�s���{������               �m�r�ėp�l
        Dim w_AttendCompFlg_F As ADODB.Field                    '��u�σt���O               �m�r���C�\���e
        Dim w_AttendLecrep_F As ADODB.Field                     '��u��                   �m�r���C�\���e
        Dim w_CostCD_F As ADODB.Field                           '��p�R�[�h                 �m�r���C�\���e
        Dim w_CostName_F As ADODB.Field                         '��p����                   �m�r�ėp�e
        Dim w_SankaFormCD_F As ADODB.Field                      '�Q���`�ԃR�[�h             �m�r���C�\���e
        Dim w_SankaFormName_F As ADODB.Field                    '�Q���`�Ԗ���               �m�r�ėp�e
        Dim w_SSBikou_F As ADODB.Field                          '���l                       �m�r���C�\���e
        Dim w_UniqueSeqNo_F As ADODB.Field                      'UNIQUESEQNO                �m�r���C�\���e
        Dim w_ApproveFlg_F As ADODB.Field                       '���F�σt���O               �m�r���C�\���e
        Dim w_SankaFlg_F As ADODB.Field                         '�Q���t���O                 �m�r���C�Q�����t�e
        Dim w_RegistFirstTimeDate_F As ADODB.Field              '����o�^����               �u�m�r���C�e
        Dim w_LastUpdTimeDate_F As ADODB.Field                  '�ŏI�X�V����               �u�m�r���C�e
        Dim w_RegistrantID_F As ADODB.Field                     '�o�^�҂h�c                 �u�m�r���C�e
        Const WC_MSTID_STUDYSANKAFORM As String = "S018"        '���C�Q���`��-�ėp�}�X�^�h�c
        Const WC_MSTID_STUDYCOST As String = "S020"             '��p�敪-�ėp�}�X�^�h�c
        Const WC_DEFAULT_TERMTO As String = "99999999"

        Try

            fncGetStudyAttend = False

            m_lngSACount = 0

            w_strSql = ""
            'Select���@�ҏW
            w_strSql = "SELECT "
            w_strSql = w_strSql & "   VST.NENDO " & vbCrLf
            w_strSql = w_strSql & " , VST.STUDYIDX " & vbCrLf
            w_strSql = w_strSql & " , VST.OUTINFLG " & vbCrLf
            w_strSql = w_strSql & " , VST.STUDYCD " & vbCrLf
            w_strSql = w_strSql & " , VST.STUDYNAME " & vbCrLf
            w_strSql = w_strSql & " , VST.STUDYSECNAME " & vbCrLf
            w_strSql = w_strSql & " , VST.KANA " & vbCrLf
            w_strSql = w_strSql & " , VST.KINDCD " & vbCrLf
            w_strSql = w_strSql & " , VST.KINDNAME " & vbCrLf
            w_strSql = w_strSql & " , VST.SPONSORCD " & vbCrLf
            w_strSql = w_strSql & " , VST.SPONSORNAME " & vbCrLf
            w_strSql = w_strSql & " , VST.THEME " & vbCrLf
            w_strSql = w_strSql & " , VST.LECTURER " & vbCrLf
            w_strSql = w_strSql & " , VST.HALL " & vbCrLf
            w_strSql = w_strSql & " , VST.SANKACOND " & vbCrLf
            w_strSql = w_strSql & " , VST.SANKANINZU " & vbCrLf
            w_strSql = w_strSql & " , VST.REPORTS " & vbCrLf
            w_strSql = w_strSql & " , VST.BIKOU                     AS VIEWBIKOU " & vbCrLf
            w_strSql = w_strSql & " , VST.URL " & vbCrLf
            w_strSql = w_strSql & " , VST.NECESSARYVALUATIONLEVELCD " & vbCrLf
            w_strSql = w_strSql & " , VST.VALUATIONLEVELNAME " & vbCrLf
            w_strSql = w_strSql & " , VST.VALUATIONLEVELSECNAME " & vbCrLf
            w_strSql = w_strSql & " , VST.VALUATIONLEVELMARK " & vbCrLf
            w_strSql = w_strSql & " , VST.ACCEPTFROMDAY " & vbCrLf
            w_strSql = w_strSql & " , VST.ACCEPTTODAY " & vbCrLf
            w_strSql = w_strSql & " , VST.ACCEPTFLG " & vbCrLf
            w_strSql = w_strSql & " , VST.NENDOPLANKBN " & vbCrLf
            w_strSql = w_strSql & " , VST.KINMUDEPTCD " & vbCrLf
            w_strSql = w_strSql & " , VST.KINMUDEPTNAME " & vbCrLf
            w_strSql = w_strSql & " , VST.ALLDAYSNECESSARYFLG " & vbCrLf
            w_strSql = w_strSql & " , VST.INDEPENDENTFLG " & vbCrLf
            w_strSql = w_strSql & " , SD1.DATEIDX " & vbCrLf
            w_strSql = w_strSql & " , SD1.DATEAPPOFLG " & vbCrLf
            w_strSql = w_strSql & " , SD1.DATEFROM " & vbCrLf
            w_strSql = w_strSql & " , SD1.DATETO " & vbCrLf
            w_strSql = w_strSql & " , SD1.JAPANAREACD " & vbCrLf
            w_strSql = w_strSql & " , HMJ.NAME                      AS JAPANAREANAME " & vbCrLf
            w_strSql = w_strSql & " , SDA.ATTENDCOMPFLG " & vbCrLf
            w_strSql = w_strSql & " , SDA.ATTENDLECREP " & vbCrLf
            w_strSql = w_strSql & " , SDA.COSTCD " & vbCrLf
            w_strSql = w_strSql & " , HMC.NAME                      AS COSTNAME " & vbCrLf
            w_strSql = w_strSql & " , SDA.SANKAFORMCD " & vbCrLf
            w_strSql = w_strSql & " , HMS.NAME                      AS SANKAFORMNAME " & vbCrLf
            w_strSql = w_strSql & " , SDA.BIKOU                     AS APPLIBIKOU " & vbCrLf
            w_strSql = w_strSql & " , SDA.UNIQUESEQNO " & vbCrLf
            w_strSql = w_strSql & " , SDA.APPROVEFLG " & vbCrLf
            w_strSql = w_strSql & " , SDS.SANKAFLG " & vbCrLf
            w_strSql = w_strSql & " , SDA.REGISTFIRSTTIMEDATE " & vbCrLf
            w_strSql = w_strSql & " , SDA.LASTUPDTIMEDATE " & vbCrLf
            w_strSql = w_strSql & " , SDA.REGISTRANTID " & vbCrLf
            w_strSql = w_strSql & " FROM NS_STUDYAPPLI_F   SDA LEFT OUTER JOIN " & vbCrLf
            w_strSql = w_strSql & "     NS_HANYOU_M    HMS on ( " & vbCrLf
            w_strSql = w_strSql & "                               HMS.HOSPITALCD = SDA.HOSPITALCD " & vbCrLf
            w_strSql = w_strSql & "                           AND HMS.MASTERCD   = SDA.SANKAFORMCD " & vbCrLf
            w_strSql = w_strSql & "                           AND HMS.MASTERID   = '" & WC_MSTID_STUDYSANKAFORM & "'" & vbCrLf
            w_strSql = w_strSql & "                          ) " & vbCrLf
            w_strSql = w_strSql & "                           LEFT OUTER JOIN " & vbCrLf
            w_strSql = w_strSql & "     NS_HANYOU_M    HMC on ( " & vbCrLf
            w_strSql = w_strSql & "                               HMC.HOSPITALCD = SDA.HOSPITALCD " & vbCrLf
            w_strSql = w_strSql & "                           AND HMC.MASTERCD   = SDA.COSTCD " & vbCrLf
            w_strSql = w_strSql & "                           AND HMC.MASTERID   = '" & WC_MSTID_STUDYCOST & "'" & vbCrLf
            w_strSql = w_strSql & "                          ) " & vbCrLf
            w_strSql = w_strSql & " ,    NS_STUDYDATE_F SD1 LEFT OUTER JOIN " & vbCrLf
            w_strSql = w_strSql & "     NS_HANYOU_M    HMJ on ( " & vbCrLf
            w_strSql = w_strSql & "                               HMJ.HOSPITALCD = SD1.HOSPITALCD " & vbCrLf
            w_strSql = w_strSql & "                           AND HMJ.MASTERCD   = SD1.JAPANAREACD " & vbCrLf
            w_strSql = w_strSql & "                           AND HMJ.MASTERID   = '" & WC_MSTID_STUDYCOST & "'" & vbCrLf
            w_strSql = w_strSql & "                          ) " & vbCrLf
            w_strSql = w_strSql & " ,    V_NS_STUDY_F   VST " & vbCrLf
            w_strSql = w_strSql & " ,    NS_STUDYSANKADATE_F   SDS " & vbCrLf
            w_strSql = w_strSql & " ,    ( SELECT A.HOSPITALCD " & vbCrLf
            w_strSql = w_strSql & "         ,     A.NENDO " & vbCrLf
            w_strSql = w_strSql & "         ,     A.STUDYIDX " & vbCrLf
            w_strSql = w_strSql & "         ,     A.STAFFMNGID " & vbCrLf
            w_strSql = w_strSql & "         ,     MIN(B.DATEFROM) as DATEFROM " & vbCrLf
            w_strSql = w_strSql & "        FROM   NS_STUDYSANKADATE_F A " & vbCrLf
            w_strSql = w_strSql & "         ,     NS_STUDYDATE_F      B " & vbCrLf
            w_strSql = w_strSql & "        WHERE  B.HOSPITALCD = A.HOSPITALCD " & vbCrLf
            w_strSql = w_strSql & "        AND    B.NENDO      = A.NENDO " & vbCrLf
            w_strSql = w_strSql & "        AND    B.STUDYIDX   = A.STUDYIDX " & vbCrLf
            w_strSql = w_strSql & "        AND    B.DATEIDX    = A.DATEIDX " & vbCrLf
            w_strSql = w_strSql & "        AND    B.DATEFROM  <= " & IIf(m_numDateTo = 0, WC_DEFAULT_TERMTO, m_numDateTo) & "" & vbCrLf
            w_strSql = w_strSql & "        AND    B.DATETO    >= " & m_numDateFrom & "" & vbCrLf
            If m_strSankaFlg <> "" Then
                w_strSql = w_strSql & "        AND    A.SANKAFLG   = '" & m_strSankaFlg & "'" & vbCrLf
            End If
            w_strSql = w_strSql & "        GROUP BY A.HOSPITALCD , A.NENDO , A.STUDYIDX , A.STAFFMNGID " & vbCrLf
            w_strSql = w_strSql & "      ) SSS " & vbCrLf
            w_strSql = w_strSql & " WHERE   SDA.HOSPITALCD = '" & m_strHospitalCD & "'" & vbCrLf
            w_strSql = w_strSql & " AND     SDA.STAFFMNGID = '" & m_strStaffMngID & "'" & vbCrLf
            If m_numNendo <> 0 Then
                w_strSql = w_strSql & " AND     SDA.NENDO = " & m_numNendo & "" & vbCrLf
            End If
            If m_numStudyIdx <> 0 Then
                w_strSql = w_strSql & " AND     SDA.STUDYIDX = " & m_numStudyIdx & "" & vbCrLf
            End If
            If m_strAttendCompFlg <> "" Then
                w_strSql = w_strSql & " AND     SDA.ATTENDCOMPFLG = '" & m_strAttendCompFlg & "'" & vbCrLf
            End If
            If m_strApproveFlg <> "" Then
                w_strSql = w_strSql & " AND     SDA.APPROVEFLG = '" & m_strApproveFlg & "'" & vbCrLf
            End If
            w_strSql = w_strSql & " AND     COALESCE(SDA.DELFLG,'0') <> '1'" & vbCrLf
            w_strSql = w_strSql & " AND     SSS.HOSPITALCD = SDA.HOSPITALCD " & vbCrLf
            w_strSql = w_strSql & " AND     SSS.NENDO      = SDA.NENDO " & vbCrLf
            w_strSql = w_strSql & " AND     SSS.STUDYIDX   = SDA.STUDYIDX " & vbCrLf
            w_strSql = w_strSql & " AND     SSS.STAFFMNGID = SDA.STAFFMNGID " & vbCrLf
            w_strSql = w_strSql & " AND     VST.HOSPITALCD = SDA.HOSPITALCD " & vbCrLf
            w_strSql = w_strSql & " AND     VST.NENDO      = SDA.NENDO " & vbCrLf
            w_strSql = w_strSql & " AND     VST.STUDYIDX   = SDA.STUDYIDX " & vbCrLf
            If m_strOutInFlg <> "" Then
                w_strSql = w_strSql & " AND     VST.OUTINFLG   = '" & m_strOutInFlg & "'" & vbCrLf
            End If
            w_strSql = w_strSql & " AND     SDS.HOSPITALCD   = SDA.HOSPITALCD " & vbCrLf
            w_strSql = w_strSql & " AND     SDS.NENDO        = SDA.NENDO " & vbCrLf
            w_strSql = w_strSql & " AND     SDS.STUDYIDX     = SDA.STUDYIDX " & vbCrLf
            w_strSql = w_strSql & " AND     SDS.STAFFMNGID   = SDA.STAFFMNGID " & vbCrLf
            If m_strSankaFlg <> "" Then
                w_strSql = w_strSql & " AND     SDS.SANKAFLG     = '" & m_strSankaFlg & "'" & vbCrLf
            End If
            w_strSql = w_strSql & " AND     SD1.HOSPITALCD   = SDS.HOSPITALCD " & vbCrLf
            w_strSql = w_strSql & " AND     SD1.NENDO        = SDS.NENDO " & vbCrLf
            w_strSql = w_strSql & " AND     SD1.STUDYIDX     = SDS.STUDYIDX " & vbCrLf
            w_strSql = w_strSql & " AND     SD1.DATEIDX      = SDS.DATEIDX " & vbCrLf
            w_strSql = w_strSql & " AND     SD1.DATEFROM    <= " & IIf(m_numDateTo = 0, WC_DEFAULT_TERMTO, m_numDateTo) & "" & vbCrLf
            w_strSql = w_strSql & " AND     SD1.DATETO      >= " & m_numDateFrom & "" & vbCrLf

            If m_numSortFlg = 0 Then
                w_strSql = w_strSql & " ORDER BY " & vbCrLf
                w_strSql = w_strSql & "   SDA.NENDO ASC , SSS.DATEFROM ASC , SD1.DATEFROM ASC "
            ElseIf m_numSortFlg = 1 Then
                w_strSql = w_strSql & " ORDER BY " & vbCrLf
                w_strSql = w_strSql & "   SDA.NENDO DESC , SSS.DATEFROM DESC , SD1.DATEFROM DESC "
            ElseIf m_numSortFlg = "2" Then
                w_strSql = w_strSql & " ORDER BY " & vbCrLf
                w_strSql = w_strSql & "   SDA.NENDO DESC , SSS.DATEFROM DESC , SD1.DATEFROM ASC "
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�����݂��Ȃ��Ƃ�
                    ReDim g_StudyAttend(0)
                    m_lngSACount = 0
                    w_Rs.Close()
                    w_Rs = Nothing
                    Exit Function
                Else
                    '�f�[�^�����݂���Ƃ�
                    '���R�[�h�����i�[
                    .MoveLast()
                    w_intCount = .RecordCount
                    .MoveFirst()

                    '�t�B�[���h�I�u�W�F�N�g����
                    w_Nendo_F = .Fields("NENDO")                                             '�N�x                   �u�m�r���C�e
                    w_StudyIndex_F = .Fields("STUDYIDX")                                     '���C�h�m�c�d�w         �u�m�r���C�e
                    w_OutInFlg_F = .Fields("OUTINFLG")                                       '�@���O�t���O           �u�m�r���C�e
                    w_StudyCD_F = .Fields("STUDYCD")                                         '���C�R�[�h             �u�m�r���C�e
                    w_StudyName_F = .Fields("STUDYNAME")                                     '���C����               �u�m�r���C�e
                    w_StudySecName_F = .Fields("STUDYSECNAME")                               '���C����               �u�m�r���C�e
                    w_StudyKana_F = .Fields("KANA")                                          '���C�t���K�i           �u�m�r���C�e
                    w_KindCD_F = .Fields("KINDCD")                                           '��ރR�[�h             �u�m�r���C�e
                    w_KindName_F = .Fields("KINDNAME")                                       '��ޖ���               �u�m�r���C�e
                    w_SponsorCD_F = .Fields("SPONSORCD")                                     '��ÃR�[�h             �u�m�r���C�e
                    w_SponsorName_F = .Fields("SPONSORNAME")                                 '��Ö���               �u�m�r���C�e
                    w_Theme_F = .Fields("THEME")                                             '�e�[�}                 �u�m�r���C�e
                    w_Lecturer_F = .Fields("LECTURER")                                       '�u�t                   �u�m�r���C�e
                    w_Hall_F = .Fields("HALL")                                               '���E�ꏊ             �u�m�r���C�e
                    w_SankaCond_F = .Fields("SANKACOND")                                     '�Q������               �u�m�r���C�e
                    w_SankaNinzu_F = .Fields("SANKANINZU")                                   '�Q���l��               �u�m�r���C�e
                    w_Reports_F = .Fields("REPORTS")                                         '�A������               �u�m�r���C�e
                    w_Bikou_F = .Fields("VIEWBIKOU")                                         '���l                   �u�m�r���C�e
                    w_Url_F = .Fields("URL")                                                 '�t�q�k                 �u�m�r���C�e
                    w_NecessaryValuationLevelCD_F = .Fields("NECESSARYVALUATIONLEVELCD")     '�K�{�]�����x���R�[�h   �u�m�r���C�e
                    w_NecessaryValuationLevelName_F = .Fields("VALUATIONLEVELNAME")          '�K�{�]�����x������     �u�m�r���C�e
                    w_NecessaryValuationLevelSecName_F = .Fields("VALUATIONLEVELSECNAME")    '�K�{�]�����x������     �u�m�r���C�e
                    w_NecessaryValuationLevelMark_F = .Fields("VALUATIONLEVELMARK")          '�K�{�]�����x���L��     �u�m�r���C�e
                    w_AcceptFromDate_F = .Fields("ACCEPTFROMDAY")                            '��t�J�n�N����         �u�m�r���C�e
                    w_AcceptToDate_F = .Fields("ACCEPTTODAY")                                '��t�I���N����         �u�m�r���C�e
                    w_Acceptapstate_F = .Fields("ACCEPTFLG")                                 '��t�\���݃t���O       �u�m�r���C�e
                    w_NendoPlanKbn_F = .Fields("NENDOPLANKBN")                               '�N�Ԍv��敪           �u�m�r���C�e
                    w_KinmuDeptCD_F = .Fields("KINMUDEPTCD")                                 '�Ζ������R�[�h         �u�m�r���C�e
                    w_KinmuDeptName_F = .Fields("KINMUDEPTNAME")                             '�Ζ���������           �u�m�r���C�e
                    w_AllDaysNecessaryFlg_F = .Fields("ALLDAYSNECESSARYFLG")                 '�S�����K�{�t���O       �u�m�r���C�e
                    w_IndependentFlg_F = .Fields("INDEPENDENTFLG")                           '���匤�C�t���O         �u�m�r���C�e
                    w_DateIdx_F = .Fields("DATEIDX")                                         '���t�C���f�b�N�X       �m�r���C���t�e
                    w_DateAppoFlg_F = .Fields("DATEAPPOFLG")                                 '���t�w��t���O         �m�r���C���t�e
                    w_DateFrom_F = .Fields("DATEFROM")                                       '���t�J�n�N����         �m�r���C���t�e
                    w_DateTo_F = .Fields("DATETO")                                           '���t�I���N����         �m�r���C���t�e
                    w_JapanAreaCD_F = .Fields("JAPANAREACD")                                 '�s���{���R�[�h         �m�r���C���t�e
                    w_JapanAreaName_F = .Fields("JAPANAREANAME")                             '�s���{������           �m�r�ėp�l
                    w_AttendCompFlg_F = .Fields("ATTENDCOMPFLG")                             '��u�σt���O           �m�r���C�\���l
                    w_AttendLecrep_F = .Fields("ATTENDLECREP")                               '��u��               �m�r���C�\���l
                    w_CostCD_F = .Fields("COSTCD")                                           '��p�R�[�h             �m�r���C�\���l
                    w_CostName_F = .Fields("COSTNAME")                                       '��p����               �m�r�ėp�l
                    w_SankaFormCD_F = .Fields("SANKAFORMCD")                                 '�Q���`�ԃR�[�h         �m�r���C�\���l
                    w_SankaFormName_F = .Fields("SANKAFORMNAME")                             '�Q���`�Ԗ���           �m�r�ėp�l
                    w_SSBikou_F = .Fields("APPLIBIKOU")                                      '���l                   �m�r���C�\���l
                    w_UniqueSeqNo_F = .Fields("UNIQUESEQNO")                                 'UNIQUESEQNO            �m�r���C�\���l
                    w_ApproveFlg_F = .Fields("APPROVEFLG")                                   '���F�σt���O           �m�r���C�\���l
                    w_SankaFlg_F = .Fields("SANKAFLG")                                       '�Q���t���O             �m�r���C�Q�����t�l
                    w_RegistFirstTimeDate_F = .Fields("REGISTFIRSTTIMEDATE")                 '����o�^����           �u�m�r���C�e
                    w_LastUpdTimeDate_F = .Fields("LASTUPDTIMEDATE")                         '�ŏI�X�V����           �u�m�r���C�e
                    w_RegistrantID_F = .Fields("REGISTRANTID")                               '�o�^�҂h�c             �u�m�r���C�e

                    '�������z����g������
                    ReDim g_StudyAttend(w_intCount)

                    '�f�[�^���� �J��Ԃ�
                    For w_intIndex = 1 To w_intCount
                        '�N�x
                        g_StudyAttend(w_intIndex).lngNendo = Integer.Parse(General.paGetDbFieldVal(w_Nendo_F, 0))
                        '���C�h�m�c�d�w
                        g_StudyAttend(w_intIndex).lngStudyIndex = Integer.Parse(General.paGetDbFieldVal(w_StudyIndex_F, 0))
                        '�@���O�t���O
                        g_StudyAttend(w_intIndex).strOutInFlg = General.paGetDbFieldVal(w_OutInFlg_F, "")
                        '���C�R�[�h
                        g_StudyAttend(w_intIndex).strStudyCD = General.paGetDbFieldVal(w_StudyCD_F, "")
                        '���C����
                        g_StudyAttend(w_intIndex).strStudyName = General.paGetDbFieldVal(w_StudyName_F, "")
                        '���C����
                        g_StudyAttend(w_intIndex).strStudySecName = General.paGetDbFieldVal(w_StudySecName_F, "")
                        '���C�t���K�i
                        g_StudyAttend(w_intIndex).strStudyKana = General.paGetDbFieldVal(w_StudyKana_F, "")
                        '��ރR�[�h
                        g_StudyAttend(w_intIndex).strKindCD = General.paGetDbFieldVal(w_KindCD_F, "")
                        '��ޖ���
                        g_StudyAttend(w_intIndex).strKindName = General.paGetDbFieldVal(w_KindName_F, "")
                        '��ÃR�[�h
                        g_StudyAttend(w_intIndex).strSponsorCD = General.paGetDbFieldVal(w_SponsorCD_F, "")
                        '��Ö���
                        g_StudyAttend(w_intIndex).strSponsorName = General.paGetDbFieldVal(w_SponsorName_F, "")
                        '�e�[�}
                        g_StudyAttend(w_intIndex).strTheme = General.paGetDbFieldVal(w_Theme_F, "")
                        '�u�t
                        g_StudyAttend(w_intIndex).strLecturer = General.paGetDbFieldVal(w_Lecturer_F, "")
                        '���E�ꏊ
                        g_StudyAttend(w_intIndex).strHall = General.paGetDbFieldVal(w_Hall_F, "")
                        '�Q������
                        g_StudyAttend(w_intIndex).strSankaCond = General.paGetDbFieldVal(w_SankaCond_F, "")
                        '�Q���l��
                        g_StudyAttend(w_intIndex).lngSankaNinzu = Integer.Parse(General.paGetDbFieldVal(w_SankaNinzu_F, 0))
                        '�A������
                        g_StudyAttend(w_intIndex).strReports = General.paGetDbFieldVal(w_Reports_F, "")
                        '���l
                        g_StudyAttend(w_intIndex).strBikou = General.paGetDbFieldVal(w_Bikou_F, "")
                        '�t�q�k
                        g_StudyAttend(w_intIndex).strUrl = General.paGetDbFieldVal(w_Url_F, "")
                        '�K�{�]�����x���R�[�h
                        g_StudyAttend(w_intIndex).strNecessaryValuationLevelCD = General.paGetDbFieldVal(w_NecessaryValuationLevelCD_F, "")
                        '�K�{�]�����x������
                        g_StudyAttend(w_intIndex).strNecessaryValuationLevelName = General.paGetDbFieldVal(w_NecessaryValuationLevelName_F, "")
                        '�K�{�]�����x������
                        g_StudyAttend(w_intIndex).strNecessaryValuationLevelSecName = General.paGetDbFieldVal(w_NecessaryValuationLevelSecName_F, "")
                        '�K�{�]�����x���L��
                        g_StudyAttend(w_intIndex).strNecessaryValuationLevelMark = General.paGetDbFieldVal(w_NecessaryValuationLevelMark_F, "")
                        '��t�J�n�N����
                        g_StudyAttend(w_intIndex).lngAcceptFromDate = Integer.Parse(General.paGetDbFieldVal(w_AcceptFromDate_F, ""))
                        '��t�I���N����
                        g_StudyAttend(w_intIndex).lngAcceptToDate = Integer.Parse(General.paGetDbFieldVal(w_AcceptToDate_F, ""))
                        '��t�\���݃t���O
                        g_StudyAttend(w_intIndex).strAcceptapstate = General.paGetDbFieldVal(w_Acceptapstate_F, "")
                        '�N�Ԍv��敪
                        g_StudyAttend(w_intIndex).strNendoPlanKbn = General.paGetDbFieldVal(w_NendoPlanKbn_F, "")
                        '�Ζ������R�[�h
                        g_StudyAttend(w_intIndex).strKinmuDeptCD = General.paGetDbFieldVal(w_KinmuDeptCD_F, "")
                        '�Ζ���������
                        g_StudyAttend(w_intIndex).strKinmuDeptName = General.paGetDbFieldVal(w_KinmuDeptName_F, "")
                        '�S�����K�{�t���O
                        g_StudyAttend(w_intIndex).strAllDaysNecessaryFlg = General.paGetDbFieldVal(w_AllDaysNecessaryFlg_F, "")
                        '���匤�C�t���O
                        g_StudyAttend(w_intIndex).strIndependentFlg = General.paGetDbFieldVal(w_IndependentFlg_F, "")
                        '���t�C���f�b�N�X
                        g_StudyAttend(w_intIndex).lngDateIdx = Integer.Parse(General.paGetDbFieldVal(w_DateIdx_F, 0))
                        '���t�w��t���O
                        g_StudyAttend(w_intIndex).strDateAppoFlg = General.paGetDbFieldVal(w_DateAppoFlg_F, "")
                        '���t�J�n�N����
                        g_StudyAttend(w_intIndex).lngDateFrom = Integer.Parse(General.paGetDbFieldVal(w_DateFrom_F, 0))
                        '���t�I���N����
                        g_StudyAttend(w_intIndex).lngDateTo = Integer.Parse(General.paGetDbFieldVal(w_DateTo_F, 0))
                        '�s���{���R�[�h
                        g_StudyAttend(w_intIndex).strJapanAreaCD = General.paGetDbFieldVal(w_JapanAreaCD_F, "")
                        '�s���{������
                        g_StudyAttend(w_intIndex).strJapanAreaName = General.paGetDbFieldVal(w_JapanAreaName_F, "")
                        '��u�σt���O
                        g_StudyAttend(w_intIndex).strAttendCompFlg = General.paGetDbFieldVal(w_AttendCompFlg_F, "")
                        '��u��
                        g_StudyAttend(w_intIndex).strAttendLecrep = General.paGetDbFieldVal(w_AttendLecrep_F, "")
                        '��p�R�[�h
                        g_StudyAttend(w_intIndex).strCostCD = General.paGetDbFieldVal(w_CostCD_F, "")
                        '��p����
                        g_StudyAttend(w_intIndex).strCostName = General.paGetDbFieldVal(w_CostName_F, "")
                        '�Q���`�ԃR�[�h
                        g_StudyAttend(w_intIndex).strSankaFormCD = General.paGetDbFieldVal(w_SankaFormCD_F, "")
                        '�Q���`�Ԗ���
                        g_StudyAttend(w_intIndex).strSankaFormName = General.paGetDbFieldVal(w_SankaFormName_F, "")
                        '���l
                        g_StudyAttend(w_intIndex).strSSBikou = General.paGetDbFieldVal(w_SSBikou_F, "")
                        'UNIQUESEQNO
                        g_StudyAttend(w_intIndex).strUniqueSeqNo = General.paGetDbFieldVal(w_UniqueSeqNo_F, "")
                        '���F�σt���O
                        g_StudyAttend(w_intIndex).strApproveFlg = General.paGetDbFieldVal(w_ApproveFlg_F, "")
                        '�Q���t���O
                        g_StudyAttend(w_intIndex).strSankaFlg = General.paGetDbFieldVal(w_SankaFlg_F, "")

                        'Tanabe Upd 2012/11/20 Start ---�^�s��v�̏C��---******************************
                        '����o�^����
                        'g_StudyAttend(w_intIndex).dblRegistFirstTimeDate = Long.Parse(General.paGetDbFieldVal(w_RegistFirstTimeDate_F, ""))
                        g_StudyAttend(w_intIndex).dblRegistFirstTimeDate = Long.Parse(General.paGetDbFieldVal(w_RegistFirstTimeDate_F, 0))
                        '�ŏI�X�V����
                        'g_StudyAttend(w_intIndex).dblLastUpdTimeDate = Long.Parse(General.paGetDbFieldVal(w_LastUpdTimeDate_F, ""))
                        g_StudyAttend(w_intIndex).dblLastUpdTimeDate = Long.Parse(General.paGetDbFieldVal(w_LastUpdTimeDate_F, 0))
                        'Tanabe Upd 2012/11/20 End  ---------------------******************************

                        '�o�^�҂h�c
                        g_StudyAttend(w_intIndex).strRegistrantID = General.paGetDbFieldVal(w_RegistrantID_F, "")



                        '�����R�[�h�Ɉړ�
                        .MoveNext()
                    Next w_intIndex
                End If

                '�I�u�W�F�N�g ���
                .Close()
            End With

            w_Rs = Nothing

            m_lngSACount = w_intCount

            fncGetStudyAttend = True

        Catch ex As Exception

            Throw
        End Try

    End Function
    '2012/02/13 Sasaki add end-------------------------------------------------------------------------------------------------------------------------------------------
#Region "�Z���Ԑ��x"
    ''' <summary>
    ''' �Z���Ԑ��x�擾�Ҏ擾�i�ďo�j
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function mGetShortTimeIdo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetShortTimeIdo"

            mGetShortTimeIdo = False

            '�擾
            If fncGetShortTimeIdo() = False Then
                Exit Function
            End If

            '����I��
            mGetShortTimeIdo = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Z���Ԑ��x�擾�Ҏ擾�i�������j
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetShortTimeIdo() As Boolean

        Const SHORTTIMEKBN As String = "K003"
        Dim sb As New Text.StringBuilder
        Dim w_Rs As ADODB.Recordset

        Try
            fncGetShortTimeIdo = False

            sb.AppendLine("SELECT ")
            sb.AppendLine("  IDO.HOSPITALCD, ")
            sb.AppendLine("  IDO.STAFFMNGID, ")
            sb.AppendLine("  IDO.FROMDATE, ")
            sb.AppendLine("  IDO.TODATE, ")
            sb.AppendLine("  IDO.GETREASONCD, ")
            sb.AppendLine("  HAN.NAME, ")
            sb.AppendLine("  HAN.SECNAME, ")
            sb.AppendLine("  IDO.BIRTHDATE, ")
            sb.AppendLine("  IDO.REGISTFIRSTTIMEDATE, ")
            sb.AppendLine("  IDO.LASTUPDTIMEDATE, ")
            sb.AppendLine("  IDO.REGISTRANTID ")
            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals((General.gInstall_Enum.AccessType_PassThrough).ToString) Then 'ORACLE
                'ORACLE
                sb.AppendLine("FROM NS_SHORTTIMEWORKERINFO_F IDO ")
                sb.AppendLine("  ,JOIN NS_HANYOU_M HAN ")
                sb.AppendLine("WHERE  ")
                sb.AppendLine("  IDO.HOSPITALCD(+) = '" & m_strHospitalCD & "' ")
                sb.AppendLine("  AND IDO.STAFFMNGID(+) = '" & m_strStaffMngID & "' ")
                sb.AppendLine("  AND HAN.MASTERID(+) = '" & SHORTTIMEKBN & "' ") '<--�Œ�
                sb.AppendLine("  AND HAN.HOSPITALCD(+) = IDO.HOSPITALCD ")
                sb.AppendLine("  AND HAN.MASTERCD(+) = IDO.GETREASONCD ")
            Else
                '����ȊO
                sb.AppendLine("FROM NS_SHORTTIMEWORKERINFO_F IDO ")
                sb.AppendLine("LEFT OUTER JOIN NS_HANYOU_M HAN ")
                sb.AppendLine("  ON HAN.HOSPITALCD = IDO.HOSPITALCD ")
                sb.AppendLine("  AND HAN.MASTERCD = IDO.GETREASONCD ")
                sb.AppendLine("  AND HAN.MASTERID = '" & SHORTTIMEKBN & "' ") '<--�Œ�
                sb.AppendLine("WHERE  ")
                sb.AppendLine("  IDO.HOSPITALCD = '" & m_strHospitalCD & "' ")
                sb.AppendLine("  AND IDO.STAFFMNGID='" & m_strStaffMngID & "' ")

            End If

            If m_numDateFlg = 0 Then
                '�P����w��̏ꍇ
                sb.AppendLine("  AND IDO.FROMDATE <= " & m_numDateFrom & " ")
            Else
                '���Ԏw��̏ꍇ
                sb.AppendLine("  AND IDO.FROMDATE <= " & m_numDateTo & " ")
                sb.AppendLine("  AND (IDO.TODATE >= " & m_numDateFrom & " ")
                sb.AppendLine("    OR IDO.TODATE = 0 ")
                sb.AppendLine("    OR IDO.TODATE IS NULL) ")
            End If

            If m_numSortFlg = 0 Then
                '����
                sb.AppendLine("  ORDER BY IDO.FROMDATE ASC")
            ElseIf m_numSortFlg = 1 Then
                '�~��
                sb.AppendLine("  ORDER BY IDO.FROMDATE DESC")
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(sb.ToString)
            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numShortTimeKensu = 0
                    .Close()
                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numShortTimeKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_ShortTimeIdo(m_numShortTimeKensu)
                    For i As Integer = 1 To m_numShortTimeKensu
                        '�a�@�R�[�h
                        g_ShortTimeIdo(i).hospCd = General.paGetDbFieldVal(.Fields("HOSPITALCD"), "")
                        '�E���Ǘ��ԍ�
                        g_ShortTimeIdo(i).mngId = General.paGetDbFieldVal(.Fields("STAFFMNGID"), "")
                        '�J�n��
                        g_ShortTimeIdo(i).dateFrom = General.paGetDbFieldVal(.Fields("FROMDATE"), "")
                        '�I����
                        g_ShortTimeIdo(i).dateTo = General.paGetDbFieldVal(.Fields("TODATE"), "")
                        '���R�R�[�h
                        g_ShortTimeIdo(i).reasonCd = General.paGetDbFieldVal(.Fields("GETREASONCD"), "")
                        '����
                        g_ShortTimeIdo(i).name = General.paGetDbFieldVal(.Fields("NAME"), "")
                        '����
                        g_ShortTimeIdo(i).secName = General.paGetDbFieldVal(.Fields("SECNAME"), "")
                        '�o�Y��
                        g_ShortTimeIdo(i).birthDate = General.paGetDbFieldVal(.Fields("BIRTHDATE"), "")
                        '����o�^����
                        g_ShortTimeIdo(i).fstRegDate = General.paGetDbFieldVal(.Fields("REGISTFIRSTTIMEDATE"), "")
                        '�ŏI�X�V����
                        g_ShortTimeIdo(i).lstUpdDate = General.paGetDbFieldVal(.Fields("LASTUPDTIMEDATE"), "")
                        '�ŏI�o�^��
                        g_ShortTimeIdo(i).lstUserId = General.paGetDbFieldVal(.Fields("REGISTRANTID"), "")

                        .MoveNext()
                    Next
                End If
            End With

            fncGetShortTimeIdo = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Z���ԗp�C���f�b�N�X�ݒ�
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property mST_ShortTimeIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mST_ShortTimeIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numShortTimeKensu Then
                    Exit Property
                End If
                m_numShortTimeIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' �Z���Ԑ��x�擾��
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fST_ShortTimeCount() As Integer
        General.g_ErrorProc = "clsStaffIdo fST_ShortTimeCount"

        Try
            fST_ShortTimeCount = m_numShortTimeKensu
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Z���Ԑ��x�ҁE�a�@�R�[�h
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fST_HospitalCD() As String
        General.g_ErrorProc = "NSC0060C fST_HospitalCD"
        Try
            If m_numShortTimeIdx = 0 OrElse m_numShortTimeKensu = 0 Then
                fST_HospitalCD = ""
            Else
                fST_HospitalCD = g_ShortTimeIdo(m_numShortTimeIdx).hospCd
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Z���Ԑ��x�ҁE�E���Ǘ��ԍ�
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fST_StaffMngID() As String
        General.g_ErrorProc = "NSC0060C fST_StaffMngID"
        Try
            If m_numShortTimeIdx = 0 OrElse m_numShortTimeKensu = 0 Then
                fST_StaffMngID = ""
            Else
                fST_StaffMngID = g_ShortTimeIdo(m_numShortTimeIdx).mngId
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Z���Ԑ��x�ҁE�J�n��
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fST_DateFrom() As Integer
        General.g_ErrorProc = "NSC0060C fST_DateFrom"
        Try
            If m_numShortTimeIdx = 0 OrElse m_numShortTimeKensu = 0 Then
                fST_DateFrom = ""
            Else
                fST_DateFrom = g_ShortTimeIdo(m_numShortTimeIdx).dateFrom
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Z���Ԑ��x�ҁE�I����
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fST_DateTo() As Integer
        General.g_ErrorProc = "NSC0060C fST_DateTo"
        Try
            If m_numShortTimeIdx = 0 OrElse m_numShortTimeKensu = 0 Then
                fST_DateTo = ""
            Else
                fST_DateTo = g_ShortTimeIdo(m_numShortTimeIdx).dateTo
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Z���Ԑ��x�ҁE���R�R�[�h
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fST_ReasonCd() As String
        General.g_ErrorProc = "NSC0060C fST_ReasonCd"
        Try
            If m_numShortTimeIdx = 0 OrElse m_numShortTimeKensu = 0 Then
                fST_ReasonCd = ""
            Else
                fST_ReasonCd = g_ShortTimeIdo(m_numShortTimeIdx).reasonCd
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Z���Ԑ��x�ҁE����
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fST_Name() As String
        General.g_ErrorProc = "NSC0060C fST_Name"
        Try
            If m_numShortTimeIdx = 0 OrElse m_numShortTimeKensu = 0 Then
                fST_Name = ""
            Else
                fST_Name = g_ShortTimeIdo(m_numShortTimeIdx).name
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Z���Ԑ��x�ҁE����
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fST_SecName() As String
        General.g_ErrorProc = "NSC0060C fST_SecName"
        Try
            If m_numShortTimeIdx = 0 OrElse m_numShortTimeKensu = 0 Then
                fST_SecName = ""
            Else
                fST_SecName = g_ShortTimeIdo(m_numShortTimeIdx).secName
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Z���Ԑ��x�ҁE�o�Y��
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fST_BirthDate() As String
        General.g_ErrorProc = "NSC0060C fST_SecName"
        Try
            If m_numShortTimeIdx = 0 OrElse m_numShortTimeKensu = 0 Then
                fST_BirthDate = ""
            Else
                fST_BirthDate = g_ShortTimeIdo(m_numShortTimeIdx).birthDate
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Z���Ԑ��x�ҁE����o�^����
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fST_FirstTime() As String
        General.g_ErrorProc = "NSC0060C fST_FirstTime"
        Try
            If m_numShortTimeIdx = 0 OrElse m_numShortTimeKensu = 0 Then
                fST_FirstTime = ""
            Else
                fST_FirstTime = g_ShortTimeIdo(m_numShortTimeIdx).fstRegDate
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Z���Ԑ��x�ҁE�ŏI�X�V����
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fST_LastUpdTime() As String
        General.g_ErrorProc = "NSC0060C fST_LastUpdTime"
        Try
            If m_numShortTimeIdx = 0 OrElse m_numShortTimeKensu = 0 Then
                fST_LastUpdTime = ""
            Else
                fST_LastUpdTime = g_ShortTimeIdo(m_numShortTimeIdx).lstUpdDate
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �Z���Ԑ��x�ҁE�o�^��ID
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fST_RegUserID() As String
        General.g_ErrorProc = "NSC0060C fST_RegUserID"
        Try
            If m_numShortTimeIdx = 0 OrElse m_numShortTimeKensu = 0 Then
                fST_RegUserID = ""
            Else
                fST_RegUserID = g_ShortTimeIdo(m_numShortTimeIdx).lstUserId
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region "��ΐ�]"
    ''' <summary>
    ''' ��ΐ�]�Ҏ擾�i�ďo�j
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function mGetNightWorkerIdo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetNightWorkerIdo"

            mGetNightWorkerIdo = False

            '�擾
            If fncGetNightWorkerIdo() = False Then
                Exit Function
            End If

            '����I��
            mGetNightWorkerIdo = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��ΐ�]�Ҏ擾�i�������j
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetNightWorkerIdo() As Boolean

        Dim sb As New Text.StringBuilder
        Dim w_Rs As ADODB.Recordset

        Try
            fncGetNightWorkerIdo = False

            sb.AppendLine("SELECT ")
            sb.AppendLine("  HOSPITALCD, ")
            sb.AppendLine("  STAFFMNGID, ")
            sb.AppendLine("  FROMDATE, ")
            sb.AppendLine("  TODATE, ")
            sb.AppendLine("  REGISTFIRSTTIMEDATE, ")
            sb.AppendLine("  LASTUPDTIMEDATE, ")
            sb.AppendLine("  REGISTRANTID ")
            If General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, General.G_INSTALLKEY, (General.gInstall_Enum.AccessType_PassThrough).ToString).Equals((General.gInstall_Enum.AccessType_PassThrough).ToString) Then 'ORACLE
                'ORACLE
                sb.AppendLine("FROM NS_NIGHTWORKERINFO_F ")
                sb.AppendLine("WHERE  ")
                sb.AppendLine("  HOSPITALCD = '" & m_strHospitalCD & "' ")
                sb.AppendLine("  AND STAFFMNGID = '" & m_strStaffMngID & "' ")
            Else
                '����ȊO
                sb.AppendLine("FROM NS_NIGHTWORKERINFO_F ")
                sb.AppendLine("WHERE  ")
                sb.AppendLine("  HOSPITALCD = '" & m_strHospitalCD & "' ")
                sb.AppendLine("  AND STAFFMNGID='" & m_strStaffMngID & "' ")
            End If

            If m_numDateFlg = 0 Then
                '�P����w��̏ꍇ
                sb.AppendLine("  AND FROMDATE <= " & m_numDateFrom & " ")
            Else
                '���Ԏw��̏ꍇ
                sb.AppendLine("  AND FROMDATE <= " & m_numDateTo & " ")
                sb.AppendLine("  AND (TODATE >= " & m_numDateFrom & " ")
                sb.AppendLine("    OR TODATE = 0 ")
                sb.AppendLine("    OR TODATE IS NULL) ")
            End If

            If m_numSortFlg = 0 Then
                '����
                sb.AppendLine("  ORDER BY FROMDATE ASC")
            ElseIf m_numSortFlg = 1 Then
                '�~��
                sb.AppendLine("  ORDER BY FROMDATE DESC")
            End If

            '�J�[�\���쐬
            w_Rs = General.paDBRecordSetOpen(sb.ToString)
            With w_Rs
                If .RecordCount <= 0 Then
                    '�f�[�^�Ȃ�
                    m_numNightWorkerKensu = 0
                    .Close()
                    w_Rs = Nothing
                    Exit Function
                Else
                    .MoveLast()
                    m_numNightWorkerKensu = .RecordCount
                    .MoveFirst()

                    ReDim g_NightWorkerIdo(m_numNightWorkerKensu)
                    For i As Integer = 1 To m_numNightWorkerKensu
                        '�a�@�R�[�h
                        g_NightWorkerIdo(i).hospCd = General.paGetDbFieldVal(.Fields("HOSPITALCD"), "")
                        '�E���Ǘ��ԍ�
                        g_NightWorkerIdo(i).mngId = General.paGetDbFieldVal(.Fields("STAFFMNGID"), "")
                        '�J�n��
                        g_NightWorkerIdo(i).dateFrom = General.paGetDbFieldVal(.Fields("FROMDATE"), "")
                        '�I����
                        g_NightWorkerIdo(i).dateTo = General.paGetDbFieldVal(.Fields("TODATE"), "")
                        '����o�^����
                        g_NightWorkerIdo(i).fstRegDate = General.paGetDbFieldVal(.Fields("REGISTFIRSTTIMEDATE"), "")
                        '�ŏI�X�V����
                        g_NightWorkerIdo(i).lstUpdDate = General.paGetDbFieldVal(.Fields("LASTUPDTIMEDATE"), "")
                        '�ŏI�o�^��
                        g_NightWorkerIdo(i).lstUserId = General.paGetDbFieldVal(.Fields("REGISTRANTID"), "")

                        .MoveNext()
                    Next
                End If
            End With

            fncGetNightWorkerIdo = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��ΐ�]�җp�C���f�b�N�X�ݒ�
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property mNW_NightWorkerIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mNW_NightWorkerIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > m_numNightWorkerKensu Then
                    Exit Property
                End If
                m_numNightWorkerIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    ''' <summary>
    ''' ��ΐ�]��
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fNW_NightWorkerCount() As Integer
        General.g_ErrorProc = "clsStaffIdo fNW_NightWorkerCount"

        Try
            fNW_NightWorkerCount = m_numNightWorkerKensu
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��ΐ�]�ҁE�a�@�R�[�h
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fNW_HospitalCD() As String
        General.g_ErrorProc = "NSC0060C fNW_HospitalCD"
        Try
            If m_numNightWorkerIdx = 0 OrElse m_numNightWorkerKensu = 0 Then
                fNW_HospitalCD = ""
            Else
                fNW_HospitalCD = g_NightWorkerIdo(m_numNightWorkerIdx).hospCd
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��ΐ�]�ҁE�E���Ǘ��ԍ�
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fNW_StaffMngID() As String
        General.g_ErrorProc = "NSC0060C fNW_StaffMngID"
        Try
            If m_numNightWorkerIdx = 0 OrElse m_numNightWorkerKensu = 0 Then
                fNW_StaffMngID = ""
            Else
                fNW_StaffMngID = g_NightWorkerIdo(m_numNightWorkerIdx).mngId
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��ΐ�]�ҁE�J�n��
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fNW_DateFrom() As String
        General.g_ErrorProc = "NSC0060C fNW_DateFrom"
        Try
            If m_numNightWorkerIdx = 0 OrElse m_numNightWorkerKensu = 0 Then
                fNW_DateFrom = ""
            Else
                fNW_DateFrom = g_NightWorkerIdo(m_numNightWorkerIdx).dateFrom
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��ΐ�]�ҁE�I����
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fNW_DateTo() As String
        General.g_ErrorProc = "NSC0060C fNW_DateFrom"
        Try
            If m_numNightWorkerIdx = 0 OrElse m_numNightWorkerKensu = 0 Then
                fNW_DateTo = ""
            Else
                fNW_DateTo = g_NightWorkerIdo(m_numNightWorkerIdx).dateTo
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��ΐ�]�ҁE����o�^����
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fNW_FirstTime() As String
        General.g_ErrorProc = "NSC0060C fNW_FirstTime"
        Try
            If m_numNightWorkerIdx = 0 OrElse m_numNightWorkerKensu = 0 Then
                fNW_FirstTime = ""
            Else
                fNW_FirstTime = g_NightWorkerIdo(m_numNightWorkerIdx).fstRegDate
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��ΐ�]�ҁE�ŏI�X�V����
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fNW_LastUpdTime() As String
        General.g_ErrorProc = "NSC0060C fNW_LastUpdTime"
        Try
            If m_numNightWorkerIdx = 0 OrElse m_numNightWorkerKensu = 0 Then
                fNW_LastUpdTime = ""
            Else
                fNW_LastUpdTime = g_NightWorkerIdo(m_numNightWorkerIdx).lstUpdDate
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' ��ΐ�]�ҁE�o�^��ID
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fNW_RegUserID() As String
        General.g_ErrorProc = "NSC0060C fNW_RegUserID"
        Try
            If m_numNightWorkerIdx = 0 OrElse m_numNightWorkerKensu = 0 Then
                fNW_RegUserID = ""
            Else
                fNW_RegUserID = g_NightWorkerIdo(m_numNightWorkerIdx).lstUserId
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region
#Region "���N��ԗ���"
    ''' <summary>
    ''' ���N��ԗ����̎擾
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function mGetHealthCondHis() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetHealthCondHis"

            mGetHealthCondHis = False

            '�擾
            If fncGetHealthCondHis() = False Then
                Exit Function
            End If

            '����I��
            mGetHealthCondHis = True

        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Function fncGetHealthCondHis() As Boolean
        Dim sb As New Text.StringBuilder
        sb.AppendLine("SELECT")
        sb.AppendLine("  HEALTH.*")
        sb.AppendLine("  , HANYOU.NAME AS HEALTHCONDNAME")
        sb.AppendLine("FROM")
        sb.AppendLine("  NS_HEALTHCONDHIS_F HEALTH")
        sb.AppendLine("  LEFT OUTER JOIN NS_HANYOU_M HANYOU")
        sb.AppendLine("    ON HANYOU.HOSPITALCD = HEALTH.HOSPITALCD")
        sb.AppendLine("    AND HANYOU.MASTERID = '" & G_MASTERID_HEALTHCONDNAME & "'")
        sb.AppendLine("    AND HANYOU.MASTERCD = HEALTH.HEALTHCONDCD")
        sb.AppendLine("WHERE")
        sb.AppendLine("  HEALTH.HOSPITALCD = '" & m_strHospitalCD & "'")
        sb.AppendLine("  AND HEALTH.STAFFMNGID = '" & m_strStaffMngID & "'")
        sb.AppendLine("ORDER BY HEALTH.DISEASEDATE DESC , HEALTH.HEALTHCONDCD")
        g_HealthCondHis = General.paGetDBDataTable(sb.ToString)
        If g_HealthCondHis.Rows.Count = 0 Then
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' �{�݃R�[�h���擾����
    ''' </summary>
    ''' <returns>�{�݃R�[�h</returns>
    ''' <remarks></remarks>
    Public Function fHC_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fHC_HospitalCD"

            If m_numHealthCondHisIdx = 0 OrElse g_HealthCondHis.Rows.Count = 0 Then
                fHC_HospitalCD = ""
            Else
                fHC_HospitalCD = g_HealthCondHis.Rows(m_numHealthCondHisIdx - 1)("HOSPITALCD")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks></remarks>
    Public Function fHC_StaffmngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fHC_StaffmngID"

            If m_numHealthCondHisIdx = 0 OrElse g_HealthCondHis.Rows.Count = 0 Then
                fHC_StaffmngID = ""
            Else
                fHC_StaffmngID = g_HealthCondHis.Rows(m_numHealthCondHisIdx - 1)("STAFFMNGID")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' SEQ���擾����
    ''' </summary>
    ''' <returns>SEQ</returns>
    ''' <remarks></remarks>
    Public Function fHC_Seq() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fHC_Seq"

            If m_numHealthCondHisIdx = 0 OrElse g_HealthCondHis.Rows.Count = 0 Then
                fHC_Seq = 0
            Else
                fHC_Seq = g_HealthCondHis.Rows(m_numHealthCondHisIdx - 1)("SEQ")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' ���Ǔ����擾����
    ''' </summary>
    ''' <returns>���Ǔ�</returns>
    ''' <remarks></remarks>
    Public Function fHC_DiseaseDate() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fHC_DiseaseDate"

            If m_numHealthCondHisIdx = 0 OrElse g_HealthCondHis.Rows.Count = 0 Then
                fHC_DiseaseDate = 0
            Else
                fHC_DiseaseDate = g_HealthCondHis.Rows(m_numHealthCondHisIdx - 1)("DISEASEDATE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' ���������擾����
    ''' </summary>
    ''' <returns>������</returns>
    ''' <remarks></remarks>
    Public Function fHC_TreatDate() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fHC_TreatDate"

            If m_numHealthCondHisIdx = 0 OrElse g_HealthCondHis.Rows.Count = 0 Then
                fHC_TreatDate = 0
            Else
                fHC_TreatDate = g_HealthCondHis.Rows(m_numHealthCondHisIdx - 1)("TREATDATE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' ���N��ԃR�[�h���擾����
    ''' </summary>
    ''' <returns>���N��ԃR�[�h</returns>
    ''' <remarks></remarks>
    Public Function fHC_HealthCondCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fHC_HealthCondCD"

            If m_numHealthCondHisIdx = 0 OrElse g_HealthCondHis.Rows.Count = 0 Then
                fHC_HealthCondCD = ""
            Else
                fHC_HealthCondCD = g_HealthCondHis.Rows(m_numHealthCondHisIdx - 1)("HEALTHCONDCD")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' ���l���擾����
    ''' </summary>
    ''' <returns>���l</returns>
    ''' <remarks></remarks>
    Public Function fHC_Bikou() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fHC_Bikou"

            If m_numHealthCondHisIdx = 0 OrElse g_HealthCondHis.Rows.Count = 0 Then
                fHC_Bikou = ""
            Else
                fHC_Bikou = g_HealthCondHis.Rows(m_numHealthCondHisIdx - 1)("BIKOU")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks></remarks>
    Public Function fHC_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fHC_FirstTime"

            If m_numHealthCondHisIdx = 0 OrElse g_HealthCondHis.Rows.Count = 0 Then
                fHC_FirstTime = 0
            Else
                fHC_FirstTime = g_HealthCondHis.Rows(m_numHealthCondHisIdx - 1)("REGISTFIRSTTIMEDATE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks></remarks>
    Public Function fHC_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fHC_UpdTime"

            If m_numHealthCondHisIdx = 0 OrElse g_HealthCondHis.Rows.Count = 0 Then
                fHC_UpdTime = 0
            Else
                fHC_UpdTime = g_HealthCondHis.Rows(m_numHealthCondHisIdx - 1)("LASTUPDTIMEDATE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �o�^�҂h�c���擾����
    ''' </summary>
    ''' <returns>�o�^�҂h�c</returns>
    ''' <remarks></remarks>
    Public Function fHC_RegistrantID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fHC_RegistrantID"

            If m_numHealthCondHisIdx = 0 OrElse g_HealthCondHis.Rows.Count = 0 Then
                fHC_RegistrantID = ""
            Else
                fHC_RegistrantID = g_HealthCondHis.Rows(m_numHealthCondHisIdx - 1)("REGISTRANTID")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' ���N��Ԗ��̂��擾����
    ''' </summary>
    ''' <returns>���N��Ԗ���</returns>
    ''' <remarks></remarks>
    Public Function fHC_HealthCondName() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fHC_HealthCondName"

            If m_numHealthCondHisIdx = 0 OrElse g_HealthCondHis.Rows.Count = 0 Then
                fHC_HealthCondName = ""
            Else
                fHC_HealthCondName = g_HealthCondHis.Rows(m_numHealthCondHisIdx - 1)("HEALTHCONDNAME")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' ���N��ԗ����̌������擾����
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fHC_Count() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fHC_Count"
            If g_HealthCondHis Is Nothing OrElse g_HealthCondHis.Rows.Count = 0 Then
                Return 0
            Else
                Return g_HealthCondHis.Rows.Count
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' ���N��ԗ����������Z�b�g����
    ''' </summary>
    ''' <remarks></remarks>
    Public WriteOnly Property mHC_HealthCondHisIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mHC_HealthCondHisIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > g_HealthCondHis.Rows.Count Then
                    Exit Property
                End If
                m_numHealthCondHisIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property
#End Region
#Region "�g�a�����������"
    ''' <summary>
    ''' �g�a�����������̎擾
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function mGetHBChkHisInfo() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetHBChkHisInfo"

            mGetHBChkHisInfo = False

            '�擾
            If fncGetHBChkHisInfo() = False Then
                Exit Function
            End If

            '����I��
            mGetHBChkHisInfo = True

        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Function fncGetHBChkHisInfo() As Boolean
        Dim sb As New Text.StringBuilder
        sb.AppendLine("SELECT")
        sb.AppendLine("  *")
        sb.AppendLine("FROM")
        sb.AppendLine("  NS_HBCHKHISINFO_F")
        sb.AppendLine("WHERE")
        sb.AppendLine("  HOSPITALCD = '" & m_strHospitalCD & "'")
        sb.AppendLine("  AND STAFFMNGID = '" & m_strStaffMngID & "'")
        sb.AppendLine("ORDER BY")
        sb.AppendLine("  EXAMINEDATE DESC")

        g_HBChkHisInfo = General.paGetDBDataTable(sb.ToString)
        If g_HBChkHisInfo.Rows.Count = 0 Then
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' �{�݃R�[�h���擾����
    ''' </summary>
    ''' <returns>�{�݃R�[�h</returns>
    ''' <remarks></remarks>
    Public Function fHBI_HospitalCD() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_HospitalCD"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_HospitalCD = ""
            Else
                fHBI_HospitalCD = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("HOSPITALCD")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks></remarks>
    Public Function fHBI_StaffmngID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_StaffmngID"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_StaffmngID = ""
            Else
                fHBI_StaffmngID = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("STAFFMNGID")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �����N�������擾����
    ''' </summary>
    ''' <returns>�����N����</returns>
    ''' <remarks></remarks>
    Public Function fHBI_ExamineDate() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_ExamineDate"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_ExamineDate = ""
            Else
                fHBI_ExamineDate = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("EXAMINEDATE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' sAg�l�z���A���敪���擾����
    ''' </summary>
    ''' <returns>sAg�l�z���A���敪</returns>
    ''' <remarks></remarks>
    Public Function fHBI_SagValueYouinKbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_SagValueYouinKbn"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_SagValueYouinKbn = ""
            Else
                fHBI_SagValueYouinKbn = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("SAGVALUEYOUINKBN")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' sAg�l���擾����
    ''' </summary>
    ''' <returns>sAg�l</returns>
    ''' <remarks></remarks>
    Public Function fHBI_SagValue() As Decimal
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_SagValue"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_SagValue = ""
            Else
                fHBI_SagValue = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("SAGVALUE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' sAb�l�z���A���敪���擾����
    ''' </summary>
    ''' <returns>sAb�l�z���A���敪</returns>
    ''' <remarks></remarks>
    Public Function fHBI_SabValueYouinKbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_SabValueYouinKbn"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_SabValueYouinKbn = ""
            Else
                fHBI_SabValueYouinKbn = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("SABVALUEYOUINKBN")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' sAb�l���擾����
    ''' </summary>
    ''' <returns>sAb�l</returns>
    ''' <remarks></remarks>
    Public Function fHBI_SabValue() As Decimal
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_SabValue"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_SabValue = ""
            Else
                fHBI_SabValue = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("SABVALUE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' eAg�l�z���A���敪���擾����
    ''' </summary>
    ''' <returns>eAg�l�z���A���敪</returns>
    ''' <remarks></remarks>
    Public Function fHBI_EagValuEYouinKbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_EagValuEYouinKbn"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_EagValuEYouinKbn = ""
            Else
                fHBI_EagValuEYouinKbn = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("EAGVALUEYOUINKBN")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' eAg�l���擾����
    ''' </summary>
    ''' <returns>eAg�l</returns>
    ''' <remarks></remarks>
    Public Function fHBI_EagValue() As Decimal
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_EagValue"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_EagValue = ""
            Else
                fHBI_EagValue = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("EAGVALUE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' eAb�l�z���A���敪���擾����
    ''' </summary>
    ''' <returns>eAb�l�z���A���敪</returns>
    ''' <remarks></remarks>
    Public Function fHBI_EabValueYouinKbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_EabValueYouinKbn"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_EabValueYouinKbn = ""
            Else
                fHBI_EabValueYouinKbn = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("EABVALUEYOUINKBN")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' eAb�l���擾����
    ''' </summary>
    ''' <returns>eAb�l</returns>
    ''' <remarks></remarks>
    Public Function fHBI_EabValue() As Decimal
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_EabValue"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_EabValue = ""
            Else
                fHBI_EabValue = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("EABVALUE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' ���l���擾����
    ''' </summary>
    ''' <returns>���l</returns>
    ''' <remarks></remarks>
    Public Function fHBI_Bikou() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_Bikou"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_Bikou = ""
            Else
                fHBI_Bikou = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("BIKOU")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks></remarks>
    Public Function fHBI_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_FirstTime"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_FirstTime = 0
            Else
                fHBI_FirstTime = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("REGISTFIRSTTIMEDATE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks></remarks>
    Public Function fHBI_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_UpdTime"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_UpdTime = 0
            Else
                fHBI_UpdTime = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("LASTUPDTIMEDATE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �o�^�҂h�c���擾����
    ''' </summary>
    ''' <returns>�o�^�҂h�c</returns>
    ''' <remarks></remarks>
    Public Function fHBI_RegistrantID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_RegistrantID"

            If m_numHBChkHisInfoIdx = 0 OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                fHBI_RegistrantID = ""
            Else
                fHBI_RegistrantID = g_HBChkHisInfo.Rows(m_numHBChkHisInfoIdx - 1)("REGISTRANTID")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �g�a�����������̌������擾����
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fHBI_Count() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fHBI_Count"
            If g_HBChkHisInfo Is Nothing OrElse g_HBChkHisInfo.Rows.Count = 0 Then
                Return 0
            Else
                Return g_HBChkHisInfo.Rows.Count
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �g�a�����������������Z�b�g����
    ''' </summary>
    ''' <remarks></remarks>
    Public WriteOnly Property mHBI_HBChkHisInfoIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mHBI_HBChkHisInfoIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > g_HBChkHisInfo.Rows.Count Then
                    Exit Property
                End If
                m_numHBChkHisInfoIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property
#End Region
#Region "�����Ǘ�"

    Public Function mGetKansensyouHis() As Boolean
        Try
            General.g_ErrorProc = "clsStaffIdo mGetKansensyouHis"

            mGetKansensyouHis = False

            '�擾
            If fncGetKansensyouHis() = False Then
                Exit Function
            End If

            '����I��
            mGetKansensyouHis = True

        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Function fncGetKansensyouHis() As Boolean
        Dim sb As New Text.StringBuilder

        sb.AppendLine("SELECT")
        sb.AppendLine("  KANSENSYO.*")
        sb.AppendLine("  , HANYOU.NAME AS BYOUMEI")
        sb.AppendLine("FROM")
        sb.AppendLine("  NS_KANSENSYOUHIS_F KANSENSYO")
        sb.AppendLine("  LEFT OUTER JOIN NS_HANYOU_M HANYOU")
        sb.AppendLine("    ON KANSENSYO.BYOUMEICD = HANYOU.MASTERCD")
        sb.AppendLine("    AND HANYOU.MASTERID = '" & G_MASTERID_BYOUMEI & "'")

        sb.AppendLine("WHERE")
        sb.AppendLine("  KANSENSYO.HOSPITALCD = '" & m_strHospitalCD & "'")
        sb.AppendLine("  AND KANSENSYO.STAFFMNGID = '" & m_strStaffMngID & "'")
        sb.AppendLine("ORDER BY")
        sb.AppendLine("  KANSENSYO.REGISTDATE DESC")
        sb.AppendLine("  , KANSENSYO.BYOUMEICD")

        g_KansensyouHis = General.paGetDBDataTable(sb.ToString)
        If g_KansensyouHis.Rows.Count = 0 Then
            Return False
        End If

        Return True
    End Function
    ''' <summary>
    ''' �{�݃R�[�h���擾����
    ''' </summary>
    ''' <returns>�{�݃R�[�h</returns>
    ''' <remarks></remarks>
    Public Function fKH_HospitalCd() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_HospitalCd"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_HospitalCd = ""
            Else
                fKH_HospitalCd = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("HOSPITALCD")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �E���Ǘ��ԍ����擾����
    ''' </summary>
    ''' <returns>�E���Ǘ��ԍ�</returns>
    ''' <remarks></remarks>
    Public Function fKH_StaffmngId() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_StaffmngId"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_StaffmngId = ""
            Else
                fKH_StaffmngId = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("STAFFMNGID")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �a���R�[�h���擾����
    ''' </summary>
    ''' <returns>�a���R�[�h</returns>
    ''' <remarks></remarks>
    Public Function fKH_ByoumeiCd() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_ByoumeiCd"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_ByoumeiCd = ""
            Else
                fKH_ByoumeiCd = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("BYOUMEICD")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' SEQ���擾����
    ''' </summary>
    ''' <returns>SEQ</returns>
    ''' <remarks></remarks>
    Public Function fKH_Seq() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_Seq"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_Seq = ""
            Else
                fKH_Seq = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("SEQ")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �a�����擾����
    ''' </summary>
    ''' <returns>�a��</returns>
    ''' <remarks></remarks>
    Public Function fKH_Byoumei() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_Byoumei"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_Byoumei = ""
            Else
                fKH_Byoumei = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("BYOUMEI")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �o�^�����擾����
    ''' </summary>
    ''' <returns>�o�^��</returns>
    ''' <remarks></remarks>
    Public Function fKH_RegistDate() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_RegistDate"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_RegistDate = 0
            Else
                fKH_RegistDate = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("REGISTDATE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �I�������擾����
    ''' </summary>
    ''' <returns>�I����</returns>
    ''' <remarks></remarks>
    Public Function fKH_EndDate() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_EndDate"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_EndDate = 0
            Else
                fKH_EndDate = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("ENDDATE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �����敪���擾����
    ''' </summary>
    ''' <returns>�����敪</returns>
    ''' <remarks></remarks>
    Public Function fKH_KiouKbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_KiouKbn"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_KiouKbn = ""
            Else
                fKH_KiouKbn = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("KIOUKBN")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �R�̋敪���擾����
    ''' </summary>
    ''' <returns>�R�̋敪</returns>
    ''' <remarks></remarks>
    Public Function fKH_KoutaiKbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_KoutaiKbn"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_KoutaiKbn = ""
            Else
                fKH_KoutaiKbn = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("KOUTAIKBN")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' ���N�`���ڎ�敪���擾����
    ''' </summary>
    ''' <returns>���N�`���ڎ�敪</returns>
    ''' <remarks></remarks>
    Public Function fKH_WakuchinKbn() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_WakuchinKbn"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_WakuchinKbn = ""
            Else
                fKH_WakuchinKbn = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("WAKUCHINKBN")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' ���N�`���ڎ�����擾����
    ''' </summary>
    ''' <returns>���N�`���ڎ��</returns>
    ''' <remarks></remarks>
    Public Function fKH_WakuchinDate() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_WakuchinDate"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_WakuchinDate = ""
            Else
                fKH_WakuchinDate = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("WAKUCHINDATE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' ���l���擾����
    ''' </summary>
    ''' <returns>���l</returns>
    ''' <remarks></remarks>
    Public Function fKH_Bikou() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_Bikou"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_Bikou = ""
            Else
                fKH_Bikou = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("BIKOU")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' ����o�^�������擾����
    ''' </summary>
    ''' <returns>����o�^����</returns>
    ''' <remarks></remarks>
    Public Function fKH_FirstTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_FirstTime"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_FirstTime = 0
            Else
                fKH_FirstTime = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("REGISTFIRSTTIMEDATE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �ŏI�X�V�������擾����
    ''' </summary>
    ''' <returns>�ŏI�X�V����</returns>
    ''' <remarks></remarks>
    Public Function fKH_UpdTime() As Long
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_UpdTime"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_UpdTime = 0
            Else
                fKH_UpdTime = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("LASTUPDTIMEDATE")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �o�^�҂h�c���擾����
    ''' </summary>
    ''' <returns>�o�^�҂h�c</returns>
    ''' <remarks></remarks>
    Public Function fKH_RegistrantID() As String
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_RegistrantID"

            If m_numKansensyouHisIdx = 0 OrElse g_KansensyouHis.Rows.Count = 0 Then
                fKH_RegistrantID = ""
            Else
                fKH_RegistrantID = g_KansensyouHis.Rows(m_numKansensyouHisIdx - 1)("REGISTRANTID")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �����Ǘ��̌������擾����
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fKH_Count() As Integer
        Try
            General.g_ErrorProc = "clsStaffIdo fKH_Count"
            If g_KansensyouHis Is Nothing OrElse g_KansensyouHis.Rows.Count = 0 Then
                Return 0
            Else
                Return g_KansensyouHis.Rows.Count
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' �����Ǘ��������Z�b�g����
    ''' </summary>
    ''' <remarks></remarks>
    Public WriteOnly Property mKH_KansensyouHisIdx() As Integer
        Set(ByVal Value As Integer)
            Try
                General.g_ErrorProc = "clsStaffIdo mKH_KansensyouHisIdx"

                '�f�[�^�����Ƃ̔�r
                If Value > g_KansensyouHis.Rows.Count Then
                    Exit Property
                End If
                m_numKansensyouHisIdx = Value

                Exit Property
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property
#End Region
End Class
