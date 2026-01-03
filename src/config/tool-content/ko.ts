/**
 * Korean tool content for SEO
 * Contains detailed descriptions, how-to steps, use cases, and FAQs for all tools
 * Requirements: 3.1 - Multi-language support (Korean)
 */

import { ToolContent } from '@/types/tool';

/**
 * Korean tool content map
 */
export const toolContentKo: Record<string, ToolContent> = {
  // ==================== POPULAR TOOLS ====================
  'pdf-multi-tool': {
    title: 'PDF 멀티 도구',
    metaDescription: '올인원 PDF 편집기: 병합, 분할, 정리, 삭제, 회전, 페이지 추출을 하나의 강력한 도구로.',
    keywords: ['pdf 멀티 도구', 'pdf 편집기', 'pdf 병합', 'pdf 분할', 'pdf 정리', '올인원 pdf'],
    description: `
      <p>PDF 멀티 도구는 모든 PDF 페이지 관리 작업을 위한 포괄적인 솔루션입니다. 이 강력한 올인원 도구는 여러 PDF 작업을 단일 직관적인 인터페이스로 통합하여 시간과 노력을 절약합니다.</p>
      <p>여러 문서 병합, 큰 PDF 분할, 페이지 재정렬, 불필요한 콘텐츠 삭제, 페이지 회전, 특정 섹션 추출 등 다른 애플리케이션 간 전환 없이 모든 작업을 처리할 수 있습니다.</p>
      <p>모든 처리는 브라우저에서 직접 이루어지며, 문서의 개인 정보와 보안이 보장됩니다. 파일이 서버에 업로드되지 않습니다.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 업로드 영역에 드래그 앤 드롭하거나 클릭하여 장치에서 파일을 선택합니다.' },
      { step: 2, title: '작업 선택', description: '사용 가능한 작업 중 선택: 병합, 분할, 정리, 페이지 삭제, 회전, 빈 페이지 추가, 페이지 추출.' },
      { step: 3, title: '옵션 구성', description: '페이지 범위, 회전 각도, 병합 순서 등 선택한 작업에 특정한 설정을 조정합니다.' },
      { step: 4, title: '처리 및 다운로드', description: '처리 버튼을 클릭하고 작업이 완료되면 수정된 PDF를 다운로드합니다.' },
    ],
    useCases: [
      { title: '문서 준비', description: '불필요한 페이지 제거, 콘텐츠 재정렬, 여러 파일 결합으로 제출용 문서를 준비합니다.', icon: 'file-check' },
      { title: '보고서 조립', description: '여러 보고서 섹션을 결합하고, 표지를 추가하고, 장을 하나의 전문 문서로 정리합니다.', icon: 'book-open' },
      { title: '아카이브 관리', description: '큰 아카이브 파일을 관리 가능한 섹션으로 분할하고, 관련 페이지를 추출하고, 기록 문서를 재정리합니다.', icon: 'archive' },
    ],
    faq: [
      { question: '한 번에 몇 개의 PDF를 처리할 수 있나요?', answer: '최대 10개의 PDF 파일을 동시에 업로드하고 처리할 수 있으며, 총 최대 크기는 500MB입니다.' },
      { question: '북마크가 보존되나요?', answer: '예, PDF를 병합할 때 도구는 기존 북마크를 보존하고 선택적으로 통합된 북마크 구조로 결합할 수 있습니다.' },
      { question: '페이지 제한이 있나요?', answer: '엄격한 페이지 제한은 없습니다. 도구는 수백 페이지의 문서를 처리할 수 있지만, 매우 큰 파일은 처리 시간이 더 오래 걸릴 수 있습니다.' },
    ],
  },

  'merge-pdf': {
    title: 'PDF 병합',
    metaDescription: '여러 PDF 파일을 하나의 문서로 결합. 드래그 앤 드롭 재정렬이 가능한 무료 온라인 PDF 병합 도구.',
    keywords: ['pdf 병합', 'pdf 결합', 'pdf 합치기', 'pdf 머저', 'pdf 연결'],
    description: `<p>PDF 병합을 사용하면 여러 PDF 문서를 빠르고 쉽게 하나의 파일로 결합할 수 있습니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 파일 업로드', description: '여러 PDF 파일을 업로드 영역에 드래그 앤 드롭하거나 클릭하여 장치에서 파일을 선택합니다.' },
      { step: 2, title: '순서 정렬', description: '파일 썸네일을 드래그 앤 드롭하여 원하는 순서로 정렬합니다.' },
      { step: 3, title: '병합 및 다운로드', description: '병합 버튼을 클릭하여 모든 파일을 결합하고 병합된 PDF를 다운로드합니다.' },
    ],
    useCases: [
      { title: '보고서 결합', description: '월간 또는 분기 보고서를 하나의 연간 문서로 병합합니다.', icon: 'file-text' },
      { title: '포트폴리오 조립', description: '여러 프로젝트 문서, 인증서, 작업 샘플을 전문 포트폴리오로 결합합니다.', icon: 'briefcase' },
      { title: '청구서 통합', description: '여러 청구서나 영수증을 하나의 문서로 병합합니다.', icon: 'receipt' },
    ],
    faq: [
      { question: '몇 개의 PDF를 병합할 수 있나요?', answer: '최대 100개의 PDF 파일을 한 번에 병합할 수 있습니다.' },
      { question: '병합된 PDF는 원본 품질을 유지하나요?', answer: '예, 병합 프로세스는 모든 문서의 원본 품질을 보존합니다.' },
      { question: '암호로 보호된 PDF를 병합할 수 있나요?', answer: '암호로 보호된 PDF는 먼저 복호화해야 합니다.' },
    ],
  },

  'split-pdf': {
    title: 'PDF 분할',
    metaDescription: 'PDF 파일을 여러 문서로 분할. 특정 페이지 추출 또는 페이지 범위로 나누기.',
    keywords: ['pdf 분할', 'pdf 나누기', 'pdf 분리', '페이지 추출', 'pdf 스플리터'],
    description: `<p>PDF 분할을 사용하면 하나의 PDF 문서를 여러 개의 작은 파일로 나눌 수 있습니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 분할할 파일을 선택합니다.' },
      { step: 2, title: '분할 방법 선택', description: '페이지 범위, 특정 페이지 추출, 또는 일정 간격으로 분할을 선택합니다.' },
      { step: 3, title: '페이지 범위 정의', description: '추출할 페이지 번호 또는 범위를 입력합니다.' },
      { step: 4, title: '분할 및 다운로드', description: '분할을 클릭하여 새 PDF 파일을 만들고 다운로드합니다.' },
    ],
    useCases: [
      { title: '장 추출', description: '책이나 매뉴얼을 개별 장으로 분할합니다.', icon: 'book' },
      { title: '결합 스캔 분리', description: '일괄 스캔된 문서를 개별 파일로 나눕니다.', icon: 'copy' },
      { title: '유인물 만들기', description: '프레젠테이션에서 특정 슬라이드를 추출합니다.', icon: 'presentation' },
    ],
    faq: [
      { question: 'PDF를 개별 페이지로 분할할 수 있나요?', answer: '예, "모든 페이지 분할" 옵션을 선택하면 됩니다.' },
      { question: '분할 시 북마크는 어떻게 되나요?', answer: '추출된 페이지 범위 내의 북마크는 보존됩니다.' },
      { question: '암호로 보호된 PDF를 분할할 수 있나요?', answer: '분할 전에 PDF를 복호화해야 합니다.' },
    ],
  },

  'compress-pdf': {
    title: 'PDF 압축',
    metaDescription: '품질을 유지하면서 PDF 파일 크기 줄이기.',
    keywords: ['pdf 압축', 'pdf 크기 줄이기', 'pdf 압축기', 'pdf 축소', 'pdf 최적화'],
    description: `<p>PDF 압축은 허용 가능한 품질을 유지하면서 PDF 문서의 파일 크기를 줄입니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 압축할 문서를 선택합니다.' },
      { step: 2, title: '압축 수준 선택', description: '낮음 (최고 품질), 중간 (균형), 높음 (최소 크기) 중 선택합니다.' },
      { step: 3, title: '압축 및 다운로드', description: '압축을 클릭하여 파일 크기를 줄이고 다운로드합니다.' },
    ],
    useCases: [
      { title: '이메일 첨부', description: '이메일 첨부 제한을 충족하기 위해 PDF 크기를 줄입니다.', icon: 'mail' },
      { title: '웹 게시', description: '웹 다운로드용 PDF를 최적화합니다.', icon: 'globe' },
      { title: '저장소 최적화', description: '아카이브 문서를 압축하여 디스크 공간을 절약합니다.', icon: 'hard-drive' },
    ],
    faq: [
      { question: '파일 크기를 얼마나 줄일 수 있나요?', answer: '이미지가 많은 PDF는 50-80% 줄일 수 있습니다.' },
      { question: '압축이 텍스트 품질에 영향을 미치나요?', answer: '텍스트는 모든 압축 수준에서 선명하게 유지됩니다.' },
      { question: '여러 PDF를 한 번에 압축할 수 있나요?', answer: '예, 최대 10개의 PDF 파일을 동시에 압축할 수 있습니다.' },
    ],
  },

  'edit-pdf': {
    title: 'PDF 편집',
    metaDescription: '온라인으로 PDF 파일 편집. 문서에 텍스트, 이미지, 주석 추가.',
    keywords: ['pdf 편집', 'pdf 편집기', 'pdf 주석', 'pdf에 텍스트 추가'],
    description: `<p>PDF 편집은 PDF 문서를 수정하고 주석을 달기 위한 포괄적인 도구 세트를 제공합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 편집할 문서를 선택합니다.' },
      { step: 2, title: '편집 도구 선택', description: '텍스트, 하이라이트, 도형, 이미지, 코멘트 도구 중 선택합니다.' },
      { step: 3, title: '편집하기', description: '문서를 클릭하여 주석을 추가하고 요소를 배치합니다.' },
      { step: 4, title: '저장 및 다운로드', description: '저장을 클릭하여 변경 사항을 적용하고 다운로드합니다.' },
    ],
    useCases: [
      { title: '문서 검토', description: '협업 검토를 위해 코멘트와 마크업을 추가합니다.', icon: 'message-square' },
      { title: '양식 작성', description: '인쇄 없이 PDF 양식을 완성합니다.', icon: 'edit-3' },
      { title: '콘텐츠 수정', description: '공유 전에 민감한 정보를 제거합니다.', icon: 'eye-off' },
    ],
    faq: [
      { question: 'PDF의 원본 텍스트를 편집할 수 있나요?', answer: '이 도구는 주석과 새 콘텐츠 추가에 중점을 둡니다.' },
      { question: '편집이 영구적인가요?', answer: '주석은 평면화하여 영구적으로 만들 수 있습니다.' },
      { question: '변경 사항을 취소할 수 있나요?', answer: '예, 실행 취소/다시 실행 기능을 지원합니다.' },
    ],
  },

  'jpg-to-pdf': {
    title: 'JPG를 PDF로',
    metaDescription: 'JPG 이미지를 PDF로 변환. 여러 JPG 파일을 하나의 PDF로 결합.',
    keywords: ['jpg를 pdf로', 'jpeg를 pdf로', 'jpg 변환', '이미지를 pdf로'],
    description: `<p>JPG를 PDF로는 JPEG 이미지를 빠르고 쉽게 PDF 문서로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'JPG 이미지 업로드', description: 'JPG 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '정렬 및 구성', description: '드래그하여 이미지를 재정렬하고 페이지 크기를 선택합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 PDF를 만들고 다운로드합니다.' },
    ],
    useCases: [
      { title: '사진 앨범', description: '휴가 사진으로 PDF 앨범을 만듭니다.', icon: 'image' },
      { title: '문서 스캔', description: '휴대폰 사진을 PDF로 변환합니다.', icon: 'camera' },
      { title: '포트폴리오 만들기', description: '디자인 샘플을 PDF 포트폴리오로 컴파일합니다.', icon: 'folder' },
    ],
    faq: [
      { question: '몇 개의 이미지를 변환할 수 있나요?', answer: '최대 100개의 JPG 이미지를 변환할 수 있습니다.' },
      { question: '이미지 품질이 보존되나요?', answer: '예, 이미지는 원본 품질로 포함됩니다.' },
      { question: '다른 페이지 크기를 설정할 수 있나요?', answer: '도구는 모든 페이지에 균일한 크기를 적용합니다.' },
    ],
  },

  'sign-pdf': {
    title: 'PDF 서명',
    metaDescription: 'PDF 문서에 전자 서명 추가.',
    keywords: ['pdf 서명', '전자 서명', 'e-서명', '디지털 서명'],
    description: `<p>PDF 서명을 사용하면 PDF 문서에 전자 서명을 빠르고 안전하게 추가할 수 있습니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '서명 만들기', description: '서명을 그리거나, 입력하거나, 이미지를 업로드합니다.' },
      { step: 3, title: '배치 및 조정', description: '문서를 클릭하여 서명을 배치하고 크기를 조정합니다.' },
      { step: 4, title: '저장 및 다운로드', description: '저장을 클릭하여 서명된 PDF를 다운로드합니다.' },
    ],
    useCases: [
      { title: '계약서 서명', description: '인쇄 없이 계약서에 전자 서명합니다.', icon: 'file-signature' },
      { title: '양식 작성', description: '신청서에 서명을 추가합니다.', icon: 'clipboard' },
      { title: '승인 워크플로', description: '검토 프로세스의 일부로 문서에 서명합니다.', icon: 'check-circle' },
    ],
    faq: [
      { question: '전자 서명은 법적으로 유효한가요?', answer: '대부분의 국가에서 법적으로 인정됩니다.' },
      { question: '서명을 저장할 수 있나요?', answer: '예, 브라우저의 로컬 저장소에 저장할 수 있습니다.' },
      { question: '여러 서명을 추가할 수 있나요?', answer: '예, 필요한 만큼 서명을 추가할 수 있습니다.' },
    ],
  },

  'crop-pdf': {
    title: 'PDF 자르기',
    metaDescription: 'PDF 페이지를 잘라 여백과 원하지 않는 영역 제거.',
    keywords: ['pdf 자르기', 'pdf 트리밍', 'pdf 여백 제거'],
    description: `<p>PDF 자르기를 사용하면 PDF 페이지에서 여백과 원하지 않는 영역을 트리밍할 수 있습니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '자르기 영역 정의', description: '자르기 핸들을 드래그하여 유지할 영역을 정의합니다.' },
      { step: 3, title: '페이지에 적용', description: '모든 페이지 또는 특정 페이지에 자르기를 적용합니다.' },
      { step: 4, title: '자르기 및 다운로드', description: '자르기를 클릭하여 다운로드합니다.' },
    ],
    useCases: [
      { title: '여백 제거', description: '스캔 문서에서 과도한 여백을 트리밍합니다.', icon: 'maximize-2' },
      { title: '콘텐츠 집중', description: '특정 콘텐츠 영역을 강조합니다.', icon: 'target' },
      { title: '페이지 표준화', description: '모든 페이지를 동일한 크기로 만듭니다.', icon: 'square' },
    ],
    faq: [
      { question: '자르기가 콘텐츠를 영구적으로 제거하나요?', answer: '예, 자르기 영역 외부의 콘텐츠가 제거됩니다.' },
      { question: '다른 페이지를 다르게 자를 수 있나요?', answer: '예, 개별 페이지에 다른 설정을 적용할 수 있습니다.' },
      { question: '자르기가 텍스트 품질에 영향을 미치나요?', answer: '아니요, 남은 콘텐츠는 원본 품질을 유지합니다.' },
    ],
  },

  'extract-pages': {
    title: '페이지 추출',
    metaDescription: 'PDF 파일에서 특정 페이지 추출.',
    keywords: ['pdf 페이지 추출', 'pdf 페이지 저장', 'pdf 페이지 복사'],
    description: `<p>페이지 추출을 사용하면 PDF 문서에서 특정 페이지를 선택하여 새 파일로 저장할 수 있습니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '페이지 선택', description: '페이지 썸네일을 클릭하거나 페이지 번호를 입력합니다.' },
      { step: 3, title: '추출 및 다운로드', description: '추출을 클릭하여 새 PDF를 만들고 다운로드합니다.' },
    ],
    useCases: [
      { title: '발췌 만들기', description: '보고서에서 관련 페이지를 추출합니다.', icon: 'file-minus' },
      { title: '특정 콘텐츠 공유', description: '특정 페이지만 추출하여 공유합니다.', icon: 'share-2' },
      { title: '중요 페이지 보관', description: '주요 페이지를 추출하여 저장합니다.', icon: 'archive' },
    ],
    faq: [
      { question: '비연속 페이지를 추출할 수 있나요?', answer: '예, 모든 페이지 조합을 선택할 수 있습니다.' },
      { question: '북마크가 보존되나요?', answer: '추출된 페이지의 북마크는 보존됩니다.' },
      { question: '여러 PDF에서 페이지를 추출할 수 있나요?', answer: '이 도구는 한 번에 하나의 PDF로 작동합니다.' },
    ],
  },

  'organize-pdf': {
    title: 'PDF 정리',
    metaDescription: 'PDF 페이지 재정렬, 복제, 삭제.',
    keywords: ['pdf 정리', 'pdf 페이지 재정렬', 'pdf 재배열'],
    description: `<p>PDF 정리는 PDF 문서 내 페이지를 재배열하기 위한 직관적인 드래그 앤 드롭 인터페이스를 제공합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '페이지 재정렬', description: '페이지 썸네일을 드래그하여 재정렬합니다.' },
      { step: 3, title: '저장 및 다운로드', description: '저장을 클릭하여 재구성된 PDF를 다운로드합니다.' },
    ],
    useCases: [
      { title: '페이지 순서 수정', description: '잘못된 페이지 순서를 수정합니다.', icon: 'arrow-up-down' },
      { title: '사용자 정의 순서 만들기', description: '특정 순서로 페이지를 배열합니다.', icon: 'list' },
      { title: '원하지 않는 페이지 제거', description: '빈 페이지나 중복을 삭제합니다.', icon: 'trash-2' },
    ],
    faq: [
      { question: '페이지를 복제할 수 있나요?', answer: '예, 모든 페이지를 복제할 수 있습니다.' },
      { question: '실행 취소 기능이 있나요?', answer: '예, 변경 사항을 실행 취소할 수 있습니다.' },
      { question: '여러 PDF를 함께 정리할 수 있나요?', answer: '이 도구는 한 번에 하나의 PDF로 작동합니다.' },
    ],
  },

  'delete-pages': {
    title: '페이지 삭제',
    metaDescription: 'PDF 파일에서 원하지 않는 페이지 제거.',
    keywords: ['pdf 페이지 삭제', 'pdf 페이지 제거'],
    description: `<p>페이지 삭제를 사용하면 PDF 문서에서 원하지 않는 페이지를 빠르고 쉽게 제거할 수 있습니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '삭제할 페이지 선택', description: '페이지 썸네일을 클릭하여 삭제 대상으로 표시합니다.' },
      { step: 3, title: '삭제 및 다운로드', description: '삭제를 클릭하여 업데이트된 PDF를 다운로드합니다.' },
    ],
    useCases: [
      { title: '빈 페이지 제거', description: '실수로 포함된 빈 페이지를 제거합니다.', icon: 'file-x' },
      { title: '민감한 콘텐츠 제거', description: '기밀 정보가 포함된 페이지를 삭제합니다.', icon: 'shield' },
      { title: '문서 간소화', description: '관련 없는 페이지를 제거합니다.', icon: 'filter' },
    ],
    faq: [
      { question: '삭제된 페이지를 복구할 수 있나요?', answer: '출력 파일에서 삭제는 영구적입니다.' },
      { question: '여러 페이지를 한 번에 삭제할 수 있나요?', answer: '예, 여러 페이지를 동시에 삭제할 수 있습니다.' },
      { question: '페이지 삭제가 북마크에 영향을 미치나요?', answer: '삭제된 페이지의 북마크는 제거됩니다.' },
    ],
  },

  // ==================== EDIT & ANNOTATE ====================
  'bookmark': {
    title: '북마크 편집',
    metaDescription: 'PDF 북마크 추가, 편집, 관리. 문서 탐색 구조 생성.',
    keywords: ['pdf 북마크', '북마크 편집', '북마크 추가', 'pdf 탐색'],
    description: `<p>PDF 문서의 북마크를 생성, 수정, 정리할 수 있습니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '북마크 관리', description: '새 북마크를 추가하거나 기존 북마크를 편집합니다.' },
      { step: 3, title: '저장 및 다운로드', description: '저장을 클릭하여 업데이트된 PDF를 다운로드합니다.' },
    ],
    useCases: [
      { title: '탐색 생성', description: '긴 문서에 북마크를 추가합니다.', icon: 'navigation' },
      { title: '장 정리', description: '계층적 북마크 구조를 만듭니다.', icon: 'book-open' },
      { title: '접근성 향상', description: '문서를 더 사용하기 쉽게 만듭니다.', icon: 'accessibility' },
    ],
    faq: [
      { question: '중첩된 북마크를 만들 수 있나요?', answer: '예, 계층 구조를 만들 수 있습니다.' },
      { question: '파일에서 북마크를 가져올 수 있나요?', answer: '예, JSON 또는 텍스트 파일에서 가져올 수 있습니다.' },
      { question: '북마크가 모든 PDF 리더에서 작동하나요?', answer: '예, 표준 PDF 기능입니다.' },
    ],
  },

  'table-of-contents': {
    title: '목차',
    metaDescription: 'PDF 목차 생성. 북마크에서 클릭 가능한 탐색 생성.',
    keywords: ['pdf 목차', '목차 생성', 'pdf 인덱스'],
    description: `<p>PDF 문서의 탐색 가능한 목차 페이지를 생성합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '목차 설정', description: '북마크에서 생성하거나 사용자 정의 항목을 만듭니다.' },
      { step: 3, title: '생성 및 다운로드', description: '생성을 클릭하여 목차를 만들고 다운로드합니다.' },
    ],
    useCases: [
      { title: '학술 논문', description: '논문에 전문적인 목차를 추가합니다.', icon: 'graduation-cap' },
      { title: '비즈니스 보고서', description: '탐색 가능한 보고서를 만듭니다.', icon: 'bar-chart' },
      { title: '사용자 매뉴얼', description: '기술 문서의 목차를 생성합니다.', icon: 'book' },
    ],
    faq: [
      { question: '목차 모양을 사용자 정의할 수 있나요?', answer: '예, 스타일과 레이아웃을 선택할 수 있습니다.' },
      { question: '목차는 어디에 삽입되나요?', answer: '기본적으로 문서 시작 부분에 삽입됩니다.' },
      { question: '목차 항목이 클릭 가능한가요?', answer: '예, 각 항목은 해당 페이지로 이동하는 링크입니다.' },
    ],
  },

  'page-numbers': {
    title: '페이지 번호',
    metaDescription: 'PDF 문서에 페이지 번호 추가. 위치, 형식, 시작 번호 사용자 정의.',
    keywords: ['페이지 번호 추가', 'pdf 페이지 번호', 'pdf 페이지 매기기'],
    description: `<p>PDF 문서에 사용자 정의 가능한 페이지 번호를 추가합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '번호 매기기 설정', description: '위치, 형식, 시작 번호를 선택합니다.' },
      { step: 3, title: '적용 및 다운로드', description: '적용을 클릭하여 페이지 번호를 추가하고 다운로드합니다.' },
    ],
    useCases: [
      { title: '전문 문서', description: '보고서와 제안서에 페이지 번호를 추가합니다.', icon: 'file-text' },
      { title: '학술 논문', description: '학술 형식 요구 사항에 따라 페이지에 번호를 매깁니다.', icon: 'graduation-cap' },
      { title: '법적 문서', description: '계약서에 적절한 페이지 매기기를 추가합니다.', icon: 'scale' },
    ],
    faq: [
      { question: '첫 페이지를 건너뛸 수 있나요?', answer: '예, 번호를 매길 페이지를 지정할 수 있습니다.' },
      { question: '어떤 번호 형식을 사용할 수 있나요?', answer: '아라비아 숫자, 로마 숫자, 문자를 사용할 수 있습니다.' },
      { question: '"페이지 X/Y" 형식을 추가할 수 있나요?', answer: '예, 총 페이지 수를 포함할 수 있습니다.' },
    ],
  },

  'add-watermark': {
    title: '워터마크 추가',
    metaDescription: 'PDF 파일에 텍스트 또는 이미지 워터마크 추가.',
    keywords: ['워터마크 추가', 'pdf 워터마크', 'pdf 스탬프', 'pdf 보호'],
    description: `<p>PDF 문서에 텍스트 또는 이미지 워터마크를 배치할 수 있습니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '워터마크 생성', description: '텍스트를 입력하거나 이미지를 업로드합니다.' },
      { step: 3, title: '적용 및 다운로드', description: '적용을 클릭하여 워터마크를 추가하고 다운로드합니다.' },
    ],
    useCases: [
      { title: '문서 보호', description: '"기밀" 또는 "초안" 워터마크를 추가합니다.', icon: 'shield' },
      { title: '문서 브랜딩', description: '공식 문서에 회사 로고를 추가합니다.', icon: 'award' },
      { title: '저작권 표시', description: '지적 재산을 보호하기 위해 저작권 정보를 추가합니다.', icon: 'copyright' },
    ],
    faq: [
      { question: '이미지를 워터마크로 사용할 수 있나요?', answer: '예, PNG, JPG, SVG 이미지를 업로드할 수 있습니다.' },
      { question: '워터마크를 반투명하게 만들 수 있나요?', answer: '예, 불투명도를 조정할 수 있습니다.' },
      { question: '다른 페이지에 다른 워터마크를 적용할 수 있나요?', answer: '도구는 선택한 페이지에 동일한 워터마크를 적용합니다.' },
    ],
  },

  'header-footer': {
    title: '머리글 및 바닥글',
    metaDescription: 'PDF 문서에 머리글과 바닥글 추가.',
    keywords: ['pdf 머리글', 'pdf 바닥글', '머리글 바닥글 추가'],
    description: `<p>PDF 문서에 사용자 정의 가능한 머리글과 바닥글을 추가합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '머리글/바닥글 설정', description: '머리글과 바닥글 영역의 텍스트를 입력합니다.' },
      { step: 3, title: '적용 및 다운로드', description: '적용을 클릭하여 머리글/바닥글을 추가하고 다운로드합니다.' },
    ],
    useCases: [
      { title: '비즈니스 문서', description: '전문 문서에 회사명과 페이지 번호를 추가합니다.', icon: 'briefcase' },
      { title: '법적 문서', description: '법적 제출물에 사건 번호와 날짜를 포함합니다.', icon: 'scale' },
      { title: '학술 논문', description: '논문 제목과 저자명의 러닝 헤더를 추가합니다.', icon: 'graduation-cap' },
    ],
    faq: [
      { question: '홀수 페이지와 짝수 페이지에 다른 머리글을 설정할 수 있나요?', answer: '예, 다른 콘텐츠를 설정할 수 있습니다.' },
      { question: '현재 날짜를 포함할 수 있나요?', answer: '예, 동적 날짜 필드를 삽입할 수 있습니다.' },
      { question: '특정 페이지에서 머리글/바닥글을 건너뛸 수 있나요?', answer: '예, 건너뛸 페이지를 지정할 수 있습니다.' },
    ],
  },

  'invert-colors': {
    title: '색상 반전',
    metaDescription: '다크 모드 읽기를 위해 PDF 색상 반전.',
    keywords: ['pdf 색상 반전', 'pdf 다크 모드', '네거티브 pdf'],
    description: `<p>PDF 문서의 색상을 반전하여 네거티브 이미지 효과를 만듭니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '옵션 설정', description: '모든 콘텐츠를 반전할지 이미지를 보존할지 선택합니다.' },
      { step: 3, title: '반전 및 다운로드', description: '반전을 클릭하여 문서를 처리하고 다운로드합니다.' },
    ],
    useCases: [
      { title: '야간 읽기', description: '야간 편안한 읽기를 위한 다크 모드 버전을 만듭니다.', icon: 'moon' },
      { title: '눈의 피로 감소', description: '밝은 문서를 반전하여 눈의 피로를 줄입니다.', icon: 'eye' },
      { title: '인쇄 절약', description: '초안 인쇄 시 잉크 사용량을 줄이기 위해 문서를 반전합니다.', icon: 'printer' },
    ],
    faq: [
      { question: '이미지도 반전되나요?', answer: '기본적으로 예. 원본 이미지를 보존하도록 선택할 수 있습니다.' },
      { question: '특정 페이지만 반전할 수 있나요?', answer: '예, 반전할 페이지를 선택할 수 있습니다.' },
      { question: '반전을 되돌릴 수 있나요?', answer: '문서를 다시 반전하여 대략 원래 색상으로 돌아갈 수 있습니다.' },
    ],
  },

  'background-color': {
    title: '배경색',
    metaDescription: 'PDF 배경색 변경. 문서 페이지에 색상 배경 추가.',
    keywords: ['pdf 배경색', 'pdf 배경 변경', '색상 pdf'],
    description: `<p>PDF 페이지의 배경색을 변경하거나 추가할 수 있습니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '색상 선택', description: '색상 선택기를 사용하거나 16진수 코드를 입력합니다.' },
      { step: 3, title: '적용 및 다운로드', description: '적용을 클릭하여 배경을 추가하고 다운로드합니다.' },
    ],
    useCases: [
      { title: '가독성 향상', description: '눈의 피로를 줄이기 위해 연한 크림색 배경을 추가합니다.', icon: 'eye' },
      { title: '문서 브랜딩', description: '마케팅 자료에 브랜드 색상을 사용합니다.', icon: 'palette' },
      { title: '섹션 강조', description: '문서 섹션을 구분하기 위해 다른 배경색을 사용합니다.', icon: 'layers' },
    ],
    faq: [
      { question: '배경이 기존 콘텐츠를 덮나요?', answer: '아니요, 배경은 기존 콘텐츠 뒤에 추가됩니다.' },
      { question: '다른 페이지에 다른 색상을 사용할 수 있나요?', answer: '다른 색상을 위해 문서를 여러 번 처리해야 합니다.' },
      { question: '기존 배경을 제거할 수 있나요?', answer: '이 도구는 배경을 추가합니다. 배경 제거는 PDF 편집 도구를 사용하세요.' },
    ],
  },

  'text-color': {
    title: '텍스트 색상 변경',
    metaDescription: 'PDF 문서의 텍스트 색상 변경.',
    keywords: ['pdf 텍스트 색상 변경', 'pdf 텍스트 색상'],
    description: `<p>PDF 문서 내 텍스트의 색상을 변경할 수 있습니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '색상 선택', description: '색상 선택기를 사용하거나 16진수 코드를 입력합니다.' },
      { step: 3, title: '적용 및 다운로드', description: '적용을 클릭하여 텍스트 색상을 변경하고 다운로드합니다.' },
    ],
    useCases: [
      { title: '대비 개선', description: '배경에 대한 가독성을 향상시키기 위해 텍스트 색상을 변경합니다.', icon: 'contrast' },
      { title: '브랜드 일관성', description: '브랜드 가이드라인에 맞게 텍스트 색상을 업데이트합니다.', icon: 'palette' },
      { title: '접근성', description: '접근성 대비 요구 사항을 충족하도록 텍스트 색상을 조정합니다.', icon: 'accessibility' },
    ],
    faq: [
      { question: '모든 텍스트가 변경되나요?', answer: '예, 도구는 문서의 모든 텍스트 요소의 색상을 변경합니다.' },
      { question: '특정 텍스트만 변경할 수 있나요?', answer: '이 도구는 모든 텍스트를 변경합니다. 선택적 변경은 PDF 편집 도구를 사용하세요.' },
      { question: '서식이 지정된 텍스트(굵게, 기울임꼴)가 보존되나요?', answer: '예, 텍스트 서식은 보존됩니다. 색상만 변경됩니다.' },
    ],
  },

  'add-stamps': {
    title: '스탬프 추가',
    metaDescription: 'PDF 문서에 스탬프 추가. 승인, 검토 등을 위한 프리셋 또는 사용자 정의 스탬프 사용.',
    keywords: ['pdf 스탬프', '스탬프 추가', '승인 스탬프'],
    description: `<p>PDF 문서에 스탬프 이미지를 배치할 수 있습니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '스탬프 선택', description: '프리셋 스탬프를 선택하거나 사용자 정의 스탬프 이미지를 업로드합니다.' },
      { step: 3, title: '배치 및 적용', description: '클릭하여 스탬프를 배치하고 위치와 크기를 조정한 후 다운로드합니다.' },
    ],
    useCases: [
      { title: '문서 승인', description: '검토 워크플로에서 "승인" 또는 "거부" 스탬프를 추가합니다.', icon: 'check-circle' },
      { title: '상태 표시', description: '문서를 "초안", "최종", "기밀"로 표시합니다.', icon: 'tag' },
      { title: '품질 관리', description: '검사 또는 검토 완료를 나타내는 QC 스탬프를 추가합니다.', icon: 'clipboard-check' },
    ],
    faq: [
      { question: '어떤 프리셋 스탬프를 사용할 수 있나요?', answer: '승인, 거부, 초안, 최종, 기밀, 복사 등이 있습니다.' },
      { question: '사용자 정의 스탬프를 업로드할 수 있나요?', answer: '예, PNG 또는 JPG 이미지를 업로드할 수 있습니다.' },
      { question: '하나의 문서에 여러 스탬프를 추가할 수 있나요?', answer: '예, 여러 스탬프를 추가하고 각각 독립적으로 배치할 수 있습니다.' },
    ],
  },

  'remove-annotations': {
    title: '주석 제거',
    metaDescription: 'PDF 파일에서 주석 제거. 코멘트, 하이라이트, 마크업 삭제.',
    keywords: ['pdf 주석 제거', '코멘트 삭제', '하이라이트 제거'],
    description: `<p>PDF 문서에서 코멘트, 하이라이트, 스티커 노트 및 기타 주석을 제거합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '주석 유형 선택', description: '제거할 주석 유형을 선택합니다.' },
      { step: 3, title: '제거 및 다운로드', description: '제거를 클릭하여 주석을 삭제하고 깨끗한 PDF를 다운로드합니다.' },
    ],
    useCases: [
      { title: '문서 최종화', description: '최종 문서를 게시하기 전에 검토 코멘트를 제거합니다.', icon: 'file-check' },
      { title: '개인 정보 보호', description: '공유 전에 민감한 정보가 포함된 코멘트를 제거합니다.', icon: 'shield' },
      { title: '깨끗한 배포', description: '배포용으로 주석이 달린 문서의 깨끗한 사본을 만듭니다.', icon: 'copy' },
    ],
    faq: [
      { question: '어떤 유형의 주석을 제거할 수 있나요?', answer: '코멘트, 하이라이트, 밑줄, 취소선, 스티커 노트, 스탬프, 링크를 모두 제거할 수 있습니다.' },
      { question: '일부 주석을 유지할 수 있나요?', answer: '예, 제거할 주석 유형과 유지할 유형을 선택할 수 있습니다.' },
      { question: '이것은 되돌릴 수 있나요?', answer: '아니요, 주석 제거는 영구적입니다. 필요한 경우 원본 백업을 유지하세요.' },
    ],
  },

  'form-filler': {
    title: '양식 작성',
    metaDescription: '온라인으로 PDF 양식 작성. 인쇄 없이 대화형 PDF 양식 완성.',
    keywords: ['pdf 양식 작성', 'pdf 양식 필러', 'pdf 양식 완성'],
    description: `<p>브라우저에서 직접 대화형 PDF 양식을 작성할 수 있습니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 양식 업로드', description: 'PDF 양식을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '양식 작성', description: '양식 필드를 클릭하여 텍스트를 입력하고 체크박스를 선택합니다.' },
      { step: 3, title: '저장 및 다운로드', description: '저장을 클릭하여 작성된 양식을 다운로드합니다.' },
    ],
    useCases: [
      { title: '신청서 양식', description: '취업 신청서, 허가 신청서, 등록 양식을 완성합니다.', icon: 'clipboard' },
      { title: '세금 양식', description: '세금 문서와 재무 양식을 전자적으로 작성합니다.', icon: 'file-text' },
      { title: '계약서', description: '서명 전에 계약서 양식에 정보를 입력합니다.', icon: 'file-signature' },
    ],
    faq: [
      { question: '진행 상황을 저장할 수 있나요?', answer: '예, 부분적으로 작성된 양식을 저장하고 나중에 계속할 수 있습니다.' },
      { question: '양식 평면화란 무엇인가요?', answer: '평면화는 양식 필드를 정적 콘텐츠로 변환하여 추가 편집을 방지합니다.' },
      { question: 'XFA 양식이 지원되나요?', answer: '예, 표준 AcroForms와 XFA 양식 모두 지원됩니다.' },
    ],
  },

  'form-creator': {
    title: '양식 생성',
    metaDescription: '작성 가능한 PDF 양식 생성. 문서에 텍스트 필드, 체크박스, 드롭다운 추가.',
    keywords: ['pdf 양식 생성', 'pdf 양식 크리에이터', '작성 가능한 pdf'],
    description: `<p>정적 PDF 문서를 대화형 작성 가능한 양식으로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '양식 필드 추가', description: '도구 모음에서 필드 유형을 선택하고 클릭하여 배치합니다.' },
      { step: 3, title: '설정 및 저장', description: '필드 속성을 설정한 후 작성 가능한 PDF 양식을 저장하고 다운로드합니다.' },
    ],
    useCases: [
      { title: '신청서 양식', description: '작성 가능한 취업 신청서, 회원 양식, 등록 양식을 만듭니다.', icon: 'user-plus' },
      { title: '설문조사', description: '데이터 수집을 위한 대화형 설문조사와 질문지를 만듭니다.', icon: 'clipboard-list' },
      { title: '주문 양식', description: '수량 필드와 체크박스가 있는 제품 주문 양식을 만듭니다.', icon: 'shopping-cart' },
    ],
    faq: [
      { question: '어떤 필드 유형을 추가할 수 있나요?', answer: '텍스트 필드, 체크박스, 라디오 버튼, 드롭다운, 날짜 선택기, 서명 필드.' },
      { question: '필드를 필수로 만들 수 있나요?', answer: '예, 필드를 필수로 표시하고 유효성 검사 규칙을 추가할 수 있습니다.' },
      { question: '계산을 추가할 수 있나요?', answer: '합계와 평균 같은 기본 계산을 숫자 필드에 추가할 수 있습니다.' },
    ],
  },

  'remove-blank-pages': {
    title: '빈 페이지 제거',
    metaDescription: 'PDF 문서에서 빈 페이지 자동 감지 및 제거.',
    keywords: ['빈 페이지 제거', '빈 페이지 삭제', 'pdf 정리'],
    description: `<p>PDF 문서에서 빈 페이지를 자동으로 감지하고 제거합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '임계값 조정', description: '필요한 경우 빈 페이지 감지 임계값을 설정합니다.' },
      { step: 3, title: '제거 및 다운로드', description: '제거를 클릭하여 빈 페이지를 삭제하고 정리된 PDF를 다운로드합니다.' },
    ],
    useCases: [
      { title: '스캔 문서 정리', description: '일괄 스캔된 문서에서 빈 페이지를 제거합니다.', icon: 'scan' },
      { title: '구분자 제거', description: '병합된 문서에서 빈 구분자 페이지를 삭제합니다.', icon: 'minus' },
      { title: '파일 크기 줄이기', description: '불필요한 빈 페이지를 제거하여 문서 크기를 줄입니다.', icon: 'minimize-2' },
    ],
    faq: [
      { question: '빈 페이지 감지는 어떻게 작동하나요?', answer: '도구는 페이지 콘텐츠를 분석하고 최소한의 콘텐츠가 있는 페이지를 빈 페이지로 간주합니다.' },
      { question: '제거될 페이지를 미리 볼 수 있나요?', answer: '예, 감지된 빈 페이지는 제거 전에 검토를 위해 강조 표시됩니다.' },
      { question: '페이지에 머리글/바닥글만 있으면 어떻게 되나요?', answer: '최소한의 콘텐츠가 있는 페이지를 빈 페이지로 간주할지 결정하기 위해 임계값을 조정할 수 있습니다.' },
    ],
  },

  // ==================== CONVERT TO PDF ====================
  'image-to-pdf': {
    title: '이미지를 PDF로',
    metaDescription: '모든 이미지를 PDF로 변환. JPG, PNG, WebP, BMP, TIFF, SVG, HEIC 형식 지원.',
    keywords: ['이미지를 pdf로', '이미지 변환', '사진을 pdf로'],
    description: `<p>모든 형식의 이미지를 PDF 문서로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: '이미지 업로드', description: '지원되는 형식의 이미지를 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '정렬 및 구성', description: '이미지를 재정렬하고 페이지 크기와 방향을 선택합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 PDF를 만들고 다운로드합니다.' },
    ],
    useCases: [
      { title: '사진 컬렉션', description: '다양한 소스의 사진을 단일 PDF 앨범으로 결합합니다.', icon: 'images' },
      { title: '혼합 형식 문서', description: '다른 형식의 이미지를 통합 PDF로 변환합니다.', icon: 'file-stack' },
      { title: '아카이브 생성', description: '장기 저장을 위해 이미지 컬렉션에서 PDF 아카이브를 만듭니다.', icon: 'archive' },
    ],
    faq: [
      { question: '어떤 이미지 형식이 지원되나요?', answer: 'JPG, JPEG, PNG, WebP, BMP, TIFF, TIF, SVG, HEIC, HEIF 형식이 모두 지원됩니다.' },
      { question: '다른 이미지 형식을 혼합할 수 있나요?', answer: '예, 다른 형식의 이미지를 단일 PDF로 결합할 수 있습니다.' },
      { question: '이미지 품질이 보존되나요?', answer: '예, 압축을 선택하지 않는 한 이미지는 원본 품질로 포함됩니다.' },
    ],
  },

  'png-to-pdf': {
    title: 'PNG를 PDF로',
    metaDescription: 'PNG 이미지를 PDF로 변환. 투명도 보존 및 여러 PNG 파일 결합.',
    keywords: ['png를 pdf로', 'png 변환', 'png 변환기'],
    description: `<p>PNG 이미지를 투명도를 보존하면서 PDF 문서로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PNG 파일 업로드', description: 'PNG 이미지를 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '정렬 및 구성', description: '이미지를 재정렬하고 페이지 크기 옵션을 선택합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 PDF를 만들고 다운로드합니다.' },
    ],
    useCases: [
      { title: '그래픽 포트폴리오', description: 'PNG 그래픽과 디자인을 전문 포트폴리오로 컴파일합니다.', icon: 'palette' },
      { title: '스크린샷 문서화', description: '스크린샷을 PDF 문서로 변환합니다.', icon: 'monitor' },
      { title: '로고 컬렉션', description: '로고와 브랜드 자산의 PDF 카탈로그를 만듭니다.', icon: 'award' },
    ],
    faq: [
      { question: '투명도가 보존되나요?', answer: 'PNG 투명도는 PDF 출력에서 보존됩니다.' },
      { question: 'PNG 애니메이션은 어떻게 되나요?', answer: '애니메이션 PNG는 첫 번째 프레임을 사용하여 정적 이미지로 변환됩니다.' },
      { question: '배경색을 설정할 수 있나요?', answer: '예, 투명 영역에 대한 배경색을 선택할 수 있습니다.' },
    ],
  },

  'webp-to-pdf': {
    title: 'WebP를 PDF로',
    metaDescription: 'WebP 이미지를 PDF로 변환. 품질 보존과 함께 현대적인 이미지 형식 변환.',
    keywords: ['webp를 pdf로', 'webp 변환', 'webp 변환기'],
    description: `<p>현대적인 WebP 이미지를 PDF 문서로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'WebP 파일 업로드', description: 'WebP 이미지를 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '옵션 구성', description: '이미지를 정렬하고 페이지 크기와 방향을 선택합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 PDF를 만듭니다.' },
    ],
    useCases: [
      { title: '웹 콘텐츠 아카이빙', description: '오프라인 아카이빙을 위해 웹 이미지를 PDF로 변환합니다.', icon: 'globe' },
      { title: '인쇄 준비', description: '인쇄 목적으로 WebP 이미지를 PDF로 변환합니다.', icon: 'printer' },
      { title: '형식 표준화', description: '현대적인 WebP를 범용 호환 PDF로 변환합니다.', icon: 'file-check' },
    ],
    faq: [
      { question: 'WebP 형식이란 무엇인가요?', answer: 'WebP는 Google이 개발한 현대적인 이미지 형식으로 웹 이미지에 우수한 압축을 제공합니다.' },
      { question: '품질이 보존되나요?', answer: '예, 변환은 원본 이미지 품질을 보존합니다.' },
      { question: '애니메이션 WebP를 변환할 수 있나요?', answer: '애니메이션 WebP 파일은 정적 이미지로 변환됩니다.' },
    ],
  },

  'svg-to-pdf': {
    title: 'SVG를 PDF로',
    metaDescription: 'SVG 벡터 그래픽을 PDF로 변환. 확장성과 품질 보존.',
    keywords: ['svg를 pdf로', 'svg 변환', '벡터를 pdf로'],
    description: `<p>확장 가능한 벡터 그래픽을 벡터 품질을 보존하면서 PDF 문서로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'SVG 파일 업로드', description: 'SVG 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '설정 구성', description: '페이지 크기와 배열 옵션을 선택합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 벡터 PDF를 만듭니다.' },
    ],
    useCases: [
      { title: '로고 변환', description: 'SVG 로고를 인쇄 자료용 PDF로 변환합니다.', icon: 'award' },
      { title: '기술 도면', description: 'CAD 내보내기와 기술 일러스트레이션을 PDF로 변환합니다.', icon: 'ruler' },
      { title: '아이콘 컬렉션', description: '아이콘 세트와 그래픽의 PDF 카탈로그를 만듭니다.', icon: 'grid' },
    ],
    faq: [
      { question: '벡터 품질이 보존되나요?', answer: '예, SVG 벡터 품질은 PDF 출력에서 완전히 보존됩니다.' },
      { question: '복잡한 SVG를 변환할 수 있나요?', answer: '예, 그라디언트, 필터, 효과가 있는 복잡한 SVG가 지원됩니다.' },
      { question: '포함된 글꼴은 어떻게 되나요?', answer: 'SVG 파일에 포함된 글꼴은 PDF에서 보존됩니다.' },
    ],
  },

  'bmp-to-pdf': {
    title: 'BMP를 PDF로',
    metaDescription: 'BMP 비트맵 이미지를 PDF로 변환. 품질 보존과 함께 레거시 형식 지원.',
    keywords: ['bmp를 pdf로', 'bmp 변환', '비트맵을 pdf로'],
    description: `<p>비트맵 이미지를 PDF 문서로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'BMP 파일 업로드', description: 'BMP 이미지를 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '옵션 구성', description: '이미지를 정렬하고 페이지 설정을 선택합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 PDF를 만듭니다.' },
    ],
    useCases: [
      { title: '레거시 파일 변환', description: '오래된 BMP 파일을 현대적인 PDF 형식으로 변환합니다.', icon: 'history' },
      { title: 'Windows 스크린샷', description: 'Windows 비트맵 스크린샷을 PDF로 변환합니다.', icon: 'monitor' },
      { title: '아카이브 현대화', description: '레거시 이미지 아카이브를 PDF 형식으로 업데이트합니다.', icon: 'archive' },
    ],
    faq: [
      { question: '파일 크기가 줄어드나요?', answer: '예, BMP 파일은 PDF로 변환할 때 일반적으로 상당히 압축됩니다.' },
      { question: '품질이 보존되나요?', answer: '예, 변환 중 이미지 품질이 유지됩니다.' },
      { question: '어떤 BMP 색상 깊이가 지원되나요?', answer: '24비트 및 32비트를 포함한 모든 표준 BMP 색상 깊이가 지원됩니다.' },
    ],
  },

  'heic-to-pdf': {
    title: 'HEIC를 PDF로',
    metaDescription: 'iPhone HEIC 사진을 PDF로 변환. Apple 이미지 형식 변환이 쉬워집니다.',
    keywords: ['heic를 pdf로', 'heic 변환', 'iphone 사진을 pdf로'],
    description: `<p>Apple의 고효율 이미지 형식 사진을 PDF 문서로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'HEIC 파일 업로드', description: 'HEIC 사진을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '사진 정렬', description: '사진을 재정렬하고 페이지 설정을 선택합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 PDF를 만듭니다.' },
    ],
    useCases: [
      { title: 'iPhone 사진 앨범', description: '공유를 위해 iPhone 사진에서 PDF 앨범을 만듭니다.', icon: 'smartphone' },
      { title: '문서 스캔', description: 'iPhone 문서 스캔을 PDF 형식으로 변환합니다.', icon: 'scan' },
      { title: '크로스 플랫폼 공유', description: '범용 호환성을 위해 HEIC를 PDF로 변환합니다.', icon: 'share-2' },
    ],
    faq: [
      { question: 'HEIC 형식이란 무엇인가요?', answer: 'HEIC(고효율 이미지 컨테이너)는 JPEG보다 더 나은 압축을 제공하는 Apple의 이미지 형식입니다.' },
      { question: 'Live Photos가 지원되나요?', answer: 'Live Photos는 키 프레임을 사용하여 정적 이미지로 변환됩니다.' },
      { question: 'EXIF 데이터가 보존되나요?', answer: '사진 메타데이터는 변환 중 선택적으로 보존하거나 제거할 수 있습니다.' },
    ],
  },

  'tiff-to-pdf': {
    title: 'TIFF를 PDF로',
    metaDescription: 'TIFF 이미지를 PDF로 변환. 다중 페이지 TIFF 파일 지원 및 고품질 변환.',
    keywords: ['tiff를 pdf로', 'tiff 변환', 'tif를 pdf로'],
    description: `<p>다중 페이지 TIFF 파일을 포함한 TIFF 이미지를 PDF 문서로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'TIFF 파일 업로드', description: 'TIFF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '옵션 구성', description: '페이지 설정과 압축 옵션을 선택합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 PDF를 만듭니다.' },
    ],
    useCases: [
      { title: '스캔 문서', description: '고품질 스캔을 TIFF에서 PDF로 변환합니다.', icon: 'scan' },
      { title: '전문 그래픽', description: '배포를 위해 전문 TIFF 그래픽을 변환합니다.', icon: 'image' },
      { title: '아카이브 변환', description: 'TIFF 아카이브를 더 접근하기 쉬운 PDF 형식으로 변환합니다.', icon: 'archive' },
    ],
    faq: [
      { question: '다중 페이지 TIFF가 지원되나요?', answer: '예, 다중 페이지 TIFF 파일은 자동으로 다중 페이지 PDF로 변환됩니다.' },
      { question: '품질이 보존되나요?', answer: '예, TIFF 품질은 PDF 출력에서 완전히 보존됩니다.' },
      { question: '어떤 압축이 사용되나요?', answer: '무손실 및 손실 압축 옵션 중에서 선택할 수 있습니다.' },
    ],
  },

  'txt-to-pdf': {
    title: '텍스트를 PDF로',
    metaDescription: '일반 텍스트 파일을 PDF로 변환. 글꼴, 여백, 페이지 레이아웃 사용자 정의.',
    keywords: ['txt를 pdf로', '텍스트를 pdf로', '텍스트 파일 변환'],
    description: `<p>일반 텍스트 파일을 서식이 지정된 PDF 문서로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: '텍스트 파일 업로드', description: '.txt 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '서식 사용자 정의', description: '글꼴, 크기, 여백, 페이지 설정을 선택합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 서식이 지정된 PDF를 만듭니다.' },
    ],
    useCases: [
      { title: '코드 문서화', description: '소스 코드 파일을 문서화용 PDF로 변환합니다.', icon: 'code' },
      { title: '로그 아카이브', description: '아카이브 목적으로 로그 파일을 PDF로 변환합니다.', icon: 'file-text' },
      { title: '노트 변환', description: '일반 텍스트 노트를 서식이 지정된 PDF 문서로 변환합니다.', icon: 'sticky-note' },
    ],
    faq: [
      { question: '어떤 글꼴을 사용할 수 있나요?', answer: '코드용 고정폭 글꼴을 포함한 여러 글꼴을 사용할 수 있습니다.' },
      { question: '줄 바꿈이 자동으로 되나요?', answer: '예, 긴 줄은 페이지에 맞게 자동으로 줄 바꿈됩니다.' },
      { question: '서식을 보존할 수 있나요?', answer: '원본 텍스트의 공백과 들여쓰기가 보존됩니다.' },
    ],
  },

  'json-to-pdf': {
    title: 'JSON을 PDF로',
    metaDescription: 'JSON 파일을 서식이 지정된 PDF로 변환. 구문 강조 및 구조화된 출력.',
    keywords: ['json을 pdf로', 'json 변환', 'json 뷰어'],
    description: `<p>JSON 데이터 파일을 서식이 지정되고 읽기 쉬운 PDF 문서로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'JSON 파일 업로드', description: '.json 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '표시 구성', description: '서식 옵션과 구문 강조를 선택합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 서식이 지정된 PDF를 만듭니다.' },
    ],
    useCases: [
      { title: 'API 문서화', description: '문서화를 위해 API 응답을 PDF로 변환합니다.', icon: 'code' },
      { title: '설정 아카이브', description: '읽기 쉬운 PDF 형식으로 설정 파일을 아카이브합니다.', icon: 'settings' },
      { title: '데이터 보고서', description: 'JSON 데이터 내보내기에서 PDF 보고서를 만듭니다.', icon: 'bar-chart' },
    ],
    faq: [
      { question: '구문 강조가 포함되나요?', answer: '예, JSON 구문은 키, 값, 유형에 대해 색상으로 강조 표시됩니다.' },
      { question: '중첩된 데이터는 어떻게 처리되나요?', answer: '중첩된 객체와 배열은 가독성을 위해 적절히 들여쓰기됩니다.' },
      { question: '큰 JSON 파일은 어떻게 되나요?', answer: '큰 파일은 자동으로 여러 페이지에 걸쳐 페이지가 매겨집니다.' },
    ],
  },

  // ==================== CONVERT FROM PDF ====================
  'pdf-to-jpg': {
    title: 'PDF를 JPG로',
    metaDescription: 'PDF 페이지를 JPG 이미지로 변환. 사용자 정의 가능한 해상도로 고품질 추출.',
    keywords: ['pdf를 jpg로', 'pdf를 jpeg로', 'pdf를 이미지로 변환'],
    description: `<p>PDF 문서 페이지를 고품질 JPG 이미지로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '페이지 및 품질 선택', description: '변환할 페이지를 선택하고 품질/DPI 옵션을 설정합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 이미지를 추출하고 ZIP으로 다운로드합니다.' },
    ],
    useCases: [
      { title: '웹 게시', description: '웹사이트 사용을 위해 PDF 페이지를 이미지로 변환합니다.', icon: 'globe' },
      { title: '소셜 미디어', description: '소셜 미디어 공유를 위해 페이지를 이미지로 추출합니다.', icon: 'share-2' },
      { title: '프레젠테이션', description: '프레젠테이션용으로 PDF 슬라이드를 이미지로 변환합니다.', icon: 'presentation' },
    ],
    faq: [
      { question: '어떤 품질 설정을 사용할 수 있나요?', answer: 'DPI를 72에서 300까지, JPEG 품질을 1-100까지 설정할 수 있습니다.' },
      { question: '특정 페이지만 변환할 수 있나요?', answer: '예, 개별 페이지 또는 페이지 범위를 선택하여 변환할 수 있습니다.' },
      { question: '여러 페이지는 어떻게 처리되나요?', answer: '각 페이지는 별도의 JPG 파일이 되어 ZIP 아카이브로 다운로드됩니다.' },
    ],
  },

  'pdf-to-png': {
    title: 'PDF를 PNG로',
    metaDescription: 'PDF 페이지를 PNG 이미지로 변환. 투명도 지원과 함께 무손실 품질.',
    keywords: ['pdf를 png로', 'pdf를 png로 변환', 'pdf 이미지 추출'],
    description: `<p>PDF 문서 페이지를 무손실 압축의 고품질 PNG 이미지로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '옵션 구성', description: '페이지를 선택하고 해상도(DPI) 옵션을 설정합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 PNG 이미지를 추출합니다.' },
    ],
    useCases: [
      { title: '그래픽 추출', description: '완벽한 품질로 다이어그램과 그래픽을 추출합니다.', icon: 'image' },
      { title: '디자인 자산', description: '편집 소프트웨어용으로 PDF 디자인을 PNG로 변환합니다.', icon: 'palette' },
      { title: '문서화', description: '기술 문서용 고품질 이미지를 만듭니다.', icon: 'file-text' },
    ],
    faq: [
      { question: 'JPG 대신 PNG를 선택하는 이유는?', answer: 'PNG는 무손실 압축과 투명도 지원을 제공하여 그래픽과 텍스트에 이상적입니다.' },
      { question: '투명 배경이 지원되나요?', answer: '예, 투명도가 있는 PDF 페이지는 PNG 출력에서 보존됩니다.' },
      { question: '어떤 DPI를 사용해야 하나요?', answer: '화면 보기에는 150 DPI, 인쇄에는 300 DPI를 사용하세요.' },
    ],
  },

  'pdf-to-webp': {
    title: 'PDF를 WebP로',
    metaDescription: 'PDF 페이지를 WebP 이미지로 변환. 우수한 압축의 현대적인 형식.',
    keywords: ['pdf를 webp로', 'pdf를 webp로 변환', '현대적인 이미지 형식'],
    description: `<p>PDF 문서 페이지를 Google의 현대적인 이미지 형식인 WebP 이미지로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '품질 옵션 설정', description: '페이지를 선택하고 품질/압축 설정을 지정합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 WebP 이미지를 만듭니다.' },
    ],
    useCases: [
      { title: '웹 최적화', description: 'PDF 콘텐츠에서 웹 최적화 이미지를 만듭니다.', icon: 'globe' },
      { title: '대역폭 절약', description: '더 빠른 로딩을 위해 이미지 파일 크기를 줄입니다.', icon: 'zap' },
      { title: '현대적인 웹사이트', description: '현대적인 웹 프로젝트에 현대적인 이미지 형식을 사용합니다.', icon: 'layout' },
    ],
    faq: [
      { question: 'WebP 형식이란 무엇인가요?', answer: 'WebP는 우수한 압축을 제공하는 Google의 현대적인 이미지 형식입니다.' },
      { question: 'WebP가 널리 지원되나요?', answer: '예, 모든 현대적인 브라우저가 WebP 형식을 지원합니다.' },
      { question: 'WebP 파일이 얼마나 작나요?', answer: 'WebP 파일은 일반적으로 동등한 JPG 파일보다 25-35% 작습니다.' },
    ],
  },

  'pdf-to-bmp': {
    title: 'PDF를 BMP로',
    metaDescription: 'PDF 페이지를 BMP 비트맵 이미지로 변환. 최대 호환성을 위한 비압축 형식.',
    keywords: ['pdf를 bmp로', 'pdf를 비트맵으로 변환', '비압축 이미지'],
    description: `<p>PDF 문서 페이지를 BMP 비트맵 이미지로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '페이지 선택', description: '변환할 페이지를 선택하고 DPI를 설정합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 BMP 이미지를 만듭니다.' },
    ],
    useCases: [
      { title: '레거시 시스템', description: '오래된 소프트웨어와 호환되는 이미지를 만듭니다.', icon: 'history' },
      { title: 'Windows 애플리케이션', description: 'Windows 전용 애플리케이션용 BMP 파일을 생성합니다.', icon: 'monitor' },
      { title: '비압축 아카이브', description: 'PDF에서 비압축 이미지 아카이브를 만듭니다.', icon: 'archive' },
    ],
    faq: [
      { question: 'BMP 형식을 사용하는 이유는?', answer: 'BMP는 비압축 품질과 레거시 시스템과의 최대 호환성을 제공합니다.' },
      { question: 'BMP 파일이 더 큰가요?', answer: '예, BMP 파일은 비압축이므로 JPG나 PNG보다 상당히 큽니다.' },
      { question: '어떤 색상 깊이가 지원되나요?', answer: '24비트 및 32비트 색상 깊이가 지원됩니다.' },
    ],
  },

  'pdf-to-tiff': {
    title: 'PDF를 TIFF로',
    metaDescription: 'PDF를 TIFF 이미지로 변환. 다중 페이지 지원과 함께 전문 품질.',
    keywords: ['pdf를 tiff로', 'pdf를 tiff로 변환', '전문 이미지'],
    description: `<p>PDF 문서를 고품질 TIFF 이미지로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '출력 구성', description: '단일 또는 다중 페이지 TIFF를 선택하고 DPI를 설정합니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 TIFF 이미지를 만듭니다.' },
    ],
    useCases: [
      { title: '전문 인쇄', description: 'PDF 문서에서 인쇄 준비 TIFF 파일을 만듭니다.', icon: 'printer' },
      { title: '문서 아카이빙', description: '고품질 TIFF 형식으로 문서를 아카이브합니다.', icon: 'archive' },
      { title: '출판', description: '출판 워크플로를 위해 PDF를 TIFF로 변환합니다.', icon: 'book' },
    ],
    faq: [
      { question: '다중 페이지 TIFF를 만들 수 있나요?', answer: '예, 모든 PDF 페이지를 단일 다중 페이지 TIFF로 결합할 수 있습니다.' },
      { question: '어떤 압축 옵션을 사용할 수 있나요?', answer: 'LZW, ZIP, 무압축 옵션을 사용할 수 있습니다.' },
      { question: '인쇄에 어떤 DPI를 사용해야 하나요?', answer: '전문 인쇄에는 300 DPI 이상을 사용하세요.' },
    ],
  },

  'pdf-to-greyscale': {
    title: 'PDF를 그레이스케일로',
    metaDescription: '컬러 PDF를 그레이스케일로 변환. 파일 크기 줄이기 및 흑백 인쇄 준비.',
    keywords: ['pdf를 그레이스케일로', '그레이스케일 pdf', '흑백 pdf'],
    description: `<p>컬러 PDF 문서를 그레이스케일(흑백)로 변환합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: '컬러 PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '변환 미리보기', description: '그레이스케일 버전이 어떻게 보일지 미리 봅니다.' },
      { step: 3, title: '변환 및 다운로드', description: '변환을 클릭하여 그레이스케일 PDF를 만듭니다.' },
    ],
    useCases: [
      { title: '인쇄 절약', description: '컬러 인쇄 비용을 절약하기 위해 그레이스케일로 변환합니다.', icon: 'printer' },
      { title: '초안 문서', description: '검토용 흑백 초안을 만듭니다.', icon: 'file-text' },
      { title: '파일 크기 줄이기', description: '색상 정보를 제거하여 PDF 크기를 줄입니다.', icon: 'minimize-2' },
    ],
    faq: [
      { question: '텍스트가 읽기 쉽게 유지되나요?', answer: '예, 그레이스케일 변환 중 텍스트 선명도가 보존됩니다.' },
      { question: '파일이 얼마나 작아지나요?', answer: '파일 크기 감소는 다양하지만 색상이 많은 문서의 경우 20-50%가 될 수 있습니다.' },
      { question: '특정 페이지만 변환할 수 있나요?', answer: '예, 그레이스케일로 변환할 페이지를 선택할 수 있습니다.' },
    ],
  },

  'pdf-to-json': {
    title: 'PDF를 JSON으로',
    metaDescription: 'PDF 콘텐츠를 JSON 형식으로 추출. PDF 문서에서 구조화된 데이터 가져오기.',
    keywords: ['pdf를 json으로', 'pdf 데이터 추출', 'pdf 파서'],
    description: `<p>PDF 문서에서 콘텐츠를 구조화된 JSON 형식으로 추출합니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'PDF 파일을 드래그 앤 드롭하거나 클릭하여 선택합니다.' },
      { step: 2, title: '추출할 데이터 선택', description: '추출할 콘텐츠를 선택합니다: 텍스트, 메타데이터, 구조.' },
      { step: 3, title: '추출 및 다운로드', description: '추출을 클릭하여 JSON을 생성하고 다운로드합니다.' },
    ],
    useCases: [
      { title: '데이터 추출', description: 'PDF 문서에서 구조화된 데이터를 추출합니다.', icon: 'database' },
      { title: '문서 분석', description: 'PDF 구조와 콘텐츠를 프로그래밍 방식으로 분석합니다.', icon: 'search' },
      { title: '통합', description: 'JSON을 통해 PDF 콘텐츠를 애플리케이션으로 가져옵니다.', icon: 'plug' },
    ],
    faq: [
      { question: '어떤 데이터가 추출되나요?', answer: '텍스트 콘텐츠, 메타데이터, 페이지 크기, 글꼴, 문서 구조.' },
      { question: 'JSON 형식이 문서화되어 있나요?', answer: '예, JSON 스키마는 일관되고 잘 문서화되어 있습니다.' },
      { question: '스캔된 PDF에서 추출할 수 있나요?', answer: '스캔된 PDF는 먼저 OCR이 필요합니다. 추출 전에 OCR PDF 도구를 사용하세요.' },
    ],
  },
  'ocr-pdf': {
    title: 'OCR PDF (텍스트 인식)',
    metaDescription: '스캔한 PDF를 검색 가능한 문서로 변환. 이미지에서 텍스트를 추출하고 인식합니다.',
    keywords: ['ocr pdf', 'pdf 텍스트 인식', '스캔 pdf 검색', 'pdf 글자 추출'],
    description: `
      <p>OCR PDF 도구는 광학 문자 인식 기술을 사용하여 스캔된 문서나 이미지 기반 PDF에서 텍스트를 추출합니다. 읽기 전용 이미지를 검색과 선택이 가능한 스마트한 문서로 변환하세요.</p>
      <p>한국어, 영어 등 다양한 언어를 지원하며 원본 레이아웃을 최대한 유지하면서 검색 가능한 텍스트 레이어를 추가합니다.</p>
    `,
    howToUse: [
      { step: 1, title: '스캔된 PDF 업로드', description: '텍스트 인식이 필요한 PDF 파일을 선택하거나 드래그 앤 드롭합니다.' },
      { step: 2, title: '언어 설정', description: '문서에 포함된 언어를 선택하여 인식 정확도를 높입니다.' },
      { step: 3, title: 'OCR 실행 및 다운로드', description: '프로세스 버튼을 클릭하여 처리가 완료된 검색 가능 PDF를 저장합니다.' },
    ],
    useCases: [
      { title: '문서 디지털화', description: '종이 문서를 스캔한 파일을 디지털 데이터로 관리하고 검색합니다.', icon: 'archive' },
      { title: '텍스트 복사', description: '이미지로 된 문서에서 필요한 텍스트를 긁어내어 편집에 활용합니다.', icon: 'type' },
      { title: '접근성 향상', description: '스크린 리더가 읽을 수 없는 이미지를 텍스트로 바꾸어 접근성을 높입니다.', icon: 'accessibility' },
    ],
    faq: [
      { question: '인식률은 얼마나 되나요?', answer: '원본 스캔 품질에 따라 다르지만, 깨끗한 문서의 경우 98% 이상의 정확도를 보입니다.' },
      { question: '여러 언어가 섞여 있어도 되나요?', answer: '예, 다중 언어 설정을 통해 혼합된 언어를 인식할 수 있습니다.' },
      { question: '이미지 위치가 바뀌나요?', answer: '아니요, 원본 이미지는 그대로 유지되고 그 위에 투명한 텍스트 레이어만 입혀집니다.' },
    ],
  },

  'alternate-merge': {
    title: '교차 병합',
    metaDescription: '두 PDF의 페이지를 번갈아 가며 병합. 앞면과 뒷면 스캔 파일을 하나로 합치기에 최적입니다.',
    keywords: ['pdf 교차 병합', 'pdf 번갈아 합치기', '양면 스캔 병합'],
    description: `<p>두 개의 PDF 파일을 한 페이지씩 번갈아 가며 섞어 병합합니다. 앞면만 스캔한 파일과 뒷면만 스캔한 파일을 합칠 때 유용합니다.</p>`,
    howToUse: [
      { step: 1, title: '두 PDF 업로드', description: '앞면(홀수) PDF와 뒷면(짝수) PDF를 각각 선택합니다.' },
      { step: 2, title: '순서 및 방향 설정', description: '필요에 따라 뒷면 페이지의 순서를 역순으로 설정할 수 있습니다.' },
      { step: 3, title: '병합 실행', description: '두 파일을 섞어서 하나의 완성된 양면 문서로 다운로드합니다.' },
    ],
    useCases: [
      { title: '양면 스캔 수정', description: '단면 스캐너로 앞뒤를 따로 스캔한 결과물을 하나로 합칩니다.', icon: 'copy' },
      { title: '문서 대조', description: '두 문서의 내용을 한 페이지씩 교차로 배치하여 비교합니다.', icon: 'layers' },
    ],
    faq: [
      { question: '파일 크기가 달라도 되나요?', answer: '예, 한쪽 파일의 페이지가 먼저 끝나면 나머지 페이지는 끝에 몰아서 붙습니다.' },
      { question: '역순 스캔도 지원하나요?', answer: '예, "역순 병합" 옵션을 사용하면 거꾸로 스캔된 뒷면도 올바르게 합쳐집니다.' },
    ],
  },

  'add-attachments': {
    title: '첨부파일 추가',
    metaDescription: 'PDF 내부에 다른 파일을 포함시킵니다. 엑셀, 이미지 등 모든 형식을 첨부할 수 있습니다.',
    keywords: ['pdf 첨부파일 추가', 'pdf 파일 임베드', 'pdf 포트폴리오'],
    description: `<p>PDF 문서 안에 엑셀, 코드, 이미지 등 관련 참고 자료를 파일 형태로 직접 임베드합니다. 배포 시 여러 파일을 하나로 묶는 효과가 있습니다.</p>`,
    howToUse: [
      { step: 1, title: '메인 PDF 업로드', description: '파일을 담을 베이스가 되는 PDF를 선택합니다.' },
      { step: 2, title: '첨부할 파일 선택', description: '내부에 포함시킬 파일들을 추가합니다.' },
      { step: 3, title: '저장 및 다운로드', description: '첨부파일이 포함된 최종 PDF를 생성합니다.' },
    ],
    useCases: [
      { title: '프로젝트 제안서', description: '제안서 PDF 안에 상세 수치 엑셀 파일을 첨부하여 보냅니다.', icon: 'package' },
      { title: '증빙 서류 제출', description: '메인 문서와 함께 관련 서류들을 하나의 PDF 안에 모아 제출합니다.', icon: 'file-text' },
    ],
    faq: [
      { question: '어떤 파일 형식이 가능한가요?', answer: '형식에 제한 없이 모든 종류의 파일을 첨부할 수 있습니다.' },
      { question: '받는 사람이 파일을 볼 수 있나요?', answer: '예, 표준 PDF 뷰어의 첨부파일 탭에서 확인 및 저장이 가능합니다.' },
    ],
  },

  // ==================== SECURE & PROTECT ====================
  'encrypt-pdf': {
    title: 'PDF 암호 설정',
    metaDescription: 'PDF 파일에 비밀번호를 설정하여 문서를 보호합니다. 열기 제한 및 권한 설정을 제어하세요.',
    keywords: ['pdf 암호 설정', 'pdf 비밀번호 잠금', 'pdf 보안 설정'],
    description: `<p>강력한 암호화 알고리즘으로 PDF를 보호합니다. 문서를 열 때 필요한 비밀번호와 인쇄/복사 제한을 위한 마스터 비밀번호를 설정할 수 있습니다.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: '보호할 파일을 선택합니다.' },
      { step: 2, title: '비밀번호 입력', description: '열기 암호 또는 권한 암호를 입력합니다.' },
      { step: 3, title: '암호화 실행', description: '보안이 적용된 PDF를 다운로드합니다.' },
    ],
    useCases: [
      { title: '민감한 자료 공유', description: '급여 명세서나 계약서 등 개인정보가 포함된 파일을 안전하게 전달합니다.', icon: 'lock' },
      { title: '콘텐츠 도용 방지', description: '복사나 인쇄 기능을 비활성화하여 내용을 보호합니다.', icon: 'key' },
    ],
    faq: [
      { question: '얼마나 안전한가요?', answer: 'AES 256비트 암호화 방식을 사용하여 매우 높은 수준의 보안을 제공합니다.' },
      { question: '비밀번호를 잊어버리면 어떻게 하나요?', answer: '보안을 위해 암호화된 비밀번호는 본인 외에는 알 수 없으니 반드시 기억해 두시기 바랍니다.' },
    ],
  },

  'decrypt-pdf': {
    title: 'PDF 암호 해제',
    metaDescription: '암호로 보호된 PDF의 비밀번호를 제거하여 자유롭게 열람 및 편집할 수 있도록 합니다.',
    keywords: ['pdf 암호 해제', 'pdf 비밀번호 제거', 'pdf 잠금 풀기'],
    description: `<p>매번 입력해야 하는 번거로운 PDF 비밀번호를 제거합니다. 비밀번호를 알고 있는 경우 이를 완전히 삭제하여 일반 PDF로 만듭니다.</p>`,
    howToUse: [
      { step: 1, title: '암호화된 PDF 업로드', description: '잠금을 풀고자 하는 파일을 선택합니다.' },
      { step: 2, title: '현재 암호 입력', description: '현재 설정되어 있는 올바른 암호를 입력합니다.' },
      { step: 3, title: '해제 및 다운로드', description: '암호가 제거된 PDF를 저장합니다.' },
    ],
    useCases: [
      { title: '반복 열람 문서', description: '자주 열어봐야 하는 업무 매뉴얼의 암호를 풀어 편리하게 사용합니다.', icon: 'unlock' },
      { title: '아카이브 보관', description: '나중에 암호를 잊어버릴 상황을 대비해 보관용 파일의 암호를 미리 해제합니다.', icon: 'archive' },
    ],
    faq: [
      { question: '비밀번호를 몰라도 풀 수 있나요?', answer: '아니요, 무단 사용 방지를 위해 반드시 기존 암호를 알고 있어야 해제가 가능합니다.' },
      { question: '원본 데이터가 손상되나요?', answer: '아니요, 보안 설정만 변경될 뿐 본문 내용은 전혀 바뀌지 않습니다.' },
    ],
  },

  // ==================== OPTIMIZE & REPAIR ====================
  'repair-pdf': {
    title: 'PDF 복구',
    metaDescription: '손상되거나 열리지 않는 PDF 파일을 분석하여 복구합니다.',
    keywords: ['pdf 복구', '깨진 pdf 고치기', 'pdf 파일 수리'],
    description: `<p>파일 전송 오류나 시스템 충돌로 인해 깨진 PDF 구조를 분석하고 가능한 한 많은 데이터를 복구하여 다시 열 수 있는 상태로 만듭니다.</p>`,
    howToUse: [
      { step: 1, title: '손상된 PDF 업로드', description: '열리지 않거나 에러가 발생하는 파일을 선택합니다.' },
      { step: 2, title: '복구 프로세스 실행', description: '도구가 파일 구조를 분석하고 재구성을 시도합니다.' },
      { step: 3, title: '결과 다운로드', description: '복구된 PDF를 확인하고 저장합니다.' },
    ],
    useCases: [
      { title: '파일 에러 수정', description: '뷰어에서 열리지 않는 파일을 복구해 중요한 내용을 살려냅니다.', icon: 'refresh-cw' },
      { title: '데이터 유실 최소화', description: '일부 데이터가 깨진 문서에서 텍스트와 이미지를 최대한 추출합니다.', icon: 'wrench' },
    ],
    faq: [
      { question: '모든 파일이 복구되나요?', answer: '손상 정도에 따라 다릅니다. 구조적 오류는 대부분 고칠 수 있으나 데이터 자체가 삭제된 경우 어려울 수 있습니다.' },
      { question: '이미지가 사라졌어요.', answer: '매우 심하게 손상된 객체는 복구 과정에서 누락될 수 있습니다.' },
    ],
  },

  'pdf-to-docx': {
    title: 'PDF를 Word로 변환',
    metaDescription: 'PDF 문서를 편집 가능한 Word(DOCX) 파일로 변환합니다. 레이아웃과 서식을 완벽하게 유지하세요.',
    keywords: ['pdf word 변환', 'pdf 워드 변환', 'pdf docx 변환', 'pdf 편집 가능하게'],
    description: `
      <p>PDF 문서를 완벽하게 편집 가능한 Microsoft Word(DOCX) 파일로 변환합니다. 최신 알고리즘을 통해 텍스트 흐름, 글꼴, 표 및 이미지를 원본과 거의 동일하게 재현합니다.</p>
      <p>번거로운 타이핑 없이 PDF 내용을 Word에서 직접 수정하세요. 계약서 수정이나 보고서 재작성에 최적입니다.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: 'Word로 변환하고 싶은 파일을 선택합니다.' },
      { step: 2, title: '변환 시작', description: '문서 구조 분석이 완료될 때까지 잠시 기다립니다.' },
      { step: 3, title: 'Word 파일 다운로드', description: '생성된 DOCX 파일을 열어 자유롭게 편집합니다.' },
    ],
    useCases: [
      { title: '계약서 문구 수정', description: 'PDF로 된 계약 초안을 Word로 바꿔 조항을 수정하고 협의합니다.', icon: 'file-text' },
      { title: '이력서 업데이트', description: '오래된 PDF 이력서의 내용을 최신 경력으로 손쉽게 업데이트합니다.', icon: 'user' },
      { title: '콘텐츠 재활용', description: '보고서의 표나 텍스트를 복사해 새로운 문서를 작성할 때 활용합니다.', icon: 'copy' },
    ],
    faq: [
      { question: '서식이 깨지지는 않나요?', answer: '표준적인 문서는 레이아웃이 잘 유지됩니다. 다만, 디자인 요소가 너무 복잡하면 미세한 조정이 필요할 수 있습니다.' },
      { question: '스캔한 PDF도 되나요?', answer: '스캔본은 이미지로 인식되므로, 텍스트 편집을 원하시면 당사의 OCR 도구를 먼저 사용하시길 권장합니다.' },
      { question: '구글 문서와 호환되나요?', answer: '네, 표준 DOCX 형식이므로 구글 문서나 리브레 오피스에서도 잘 열립니다.' },
    ],
  },

  'rotate-custom': {
    title: '사용자 지정 각도 회전',
    metaDescription: 'PDF 페이지를 원하는 각도로 정밀하게 회전합니다. 기울어진 스캔 문서 교정에 적합합니다.',
    keywords: ['pdf 각도 회전', 'pdf 기울기 보정', 'pdf 수평 맞추기'],
    description: `<p>표준 90도 회전뿐만 아니라 1도 단위의 미세한 회전도 가능합니다. 잘못 들어간 스캔 문서의 수평을 맞추거나 도면의 방향을 정밀하게 조정할 때 사용하세요.</p>`,
    howToUse: [
      { step: 1, title: 'PDF 업로드', description: '회전이 필요한 페이지가 포함된 파일을 선택합니다.' },
      { step: 2, title: '각도 입력', description: '회전시키고 싶은 정확한 수치(도)를 입력합니다.' },
      { step: 3, title: '미리보기 및 적용', description: '수평이 맞는지 확인한 후 회전된 PDF를 저장합니다.' },
    ],
    useCases: [
      { title: '스캔 문서 보정', description: '비스듬하게 스캔된 종이 문서를 똑바로 세웁니다.', icon: 'scan' },
      { title: '설계도 정밀 조정', description: '도면이나 차트의 방향을 정확한 각도로 맞춥니다.', icon: 'ruler' },
    ],
    faq: [
      { question: '소수점 각도도 되나요?', answer: '현재 정수 단위 회전을 지원하며, 더 정밀한 기능은 업데이트 예정입니다.' },
      { question: '이미지가 잘리나요?', answer: '아니요, 회전된 내용이 모두 보일 수 있도록 페이지 크기가 자동으로 조정됩니다.' },
    ],
  },
};
