import difflib
import re
from collections import defaultdict

class HeaderSimilarityAnalyzer:
    """
    Advanced utility class for detecting similar or duplicate headers
    in Excel spreadsheets with multiple detection strategies.
    """
    
    def __init__(self):
        self.similarity_threshold = 0.7  # Default threshold
        self.min_word_length = 3
        self.use_string_similarity = True
        self.use_word_overlap = True
        self.detect_abbreviations = True
        self.remove_punctuation = True
        self.remove_common_prefixes = True
        
        # Common prefixes to ignore
        self.common_prefixes = ['col_', 'column_', 'data_', 'field_', 'value_']
    
    def set_similarity_threshold(self, threshold):
        """Set the similarity threshold (between 0.0 and 1.0)"""
        self.similarity_threshold = max(0.1, min(0.95, threshold))
    
    def normalize_text(self, text):
        """Normalize text for better similarity matching"""
        if text is None:
            return ""
            
        # Convert to string and lowercase
        text = str(text).lower()
        
        # Remove common prefixes if enabled
        if self.remove_common_prefixes:
            for prefix in self.common_prefixes:
                if text.startswith(prefix):
                    text = text[len(prefix):]
                    break
        
        # Remove punctuation if enabled
        if self.remove_punctuation:
            text = re.sub(r'[^\w\s]', '', text)
        
        # Replace multiple spaces with single space
        text = re.sub(r'\s+', ' ', text)
        
        # Remove leading/trailing spaces
        text = text.strip()
        
        return text
    
    def find_similar_groups(self, headers):
        """
        Find groups of similar headers using multiple techniques
        and return them grouped together
        """
        # Skip empty headers
        filtered_headers = [h for h in headers if h and str(h).strip()]
        
        # Create normalized versions for comparison
        normalized_headers = [(h, self.normalize_text(h)) for h in filtered_headers]
        
        # Find similar groups
        similar_groups = []
        processed = set()
        
        for i, (header1, norm1) in enumerate(normalized_headers):
            if header1 in processed:
                continue
                
            group = [header1]
            processed.add(header1)
            
            for j, (header2, norm2) in enumerate(normalized_headers):
                if i != j and header2 not in processed:
                    # Skip exact duplicates - they'll be handled separately
                    if norm1 == norm2:
                        continue
                    
                    # Calculate similarity with several metrics
                    similarity = self.calculate_similarity(norm1, norm2)
                    
                    if similarity >= self.similarity_threshold:
                        group.append(header2)
                        processed.add(header2)
            
            if len(group) > 1:
                similar_groups.append(group)
        
        return similar_groups
    
    def find_exact_duplicates(self, headers):
        """Find headers that are exactly the same (case-insensitive)"""
        duplicates = defaultdict(list)
        
        for header in headers:
            if header and str(header).strip():
                norm_header = self.normalize_text(header)
                duplicates[norm_header].append(header)
        
        # Return only those with multiple occurrences
        return {key: values for key, values in duplicates.items() if len(values) > 1}
    
    def find_common_word_headers(self, headers):
        """Find headers that share significant common words"""
        word_to_headers = defaultdict(list)
        significant_words = set()
        
        # Extract significant words from all headers
        for header in headers:
            if not header or not str(header).strip():
                continue
                
            words = self.normalize_text(header).split()
            # Consider words with length >= min_word_length as significant
            for word in words:
                if len(word) >= self.min_word_length:
                    significant_words.add(word)
                    word_to_headers[word].append(header)
        
        # Group headers by common significant words
        common_word_groups = []
        processed_headers = set()
        
        for word in significant_words:
            headers_with_word = word_to_headers[word]
            if len(headers_with_word) > 1:
                # Only include headers not already processed
                unprocessed = [h for h in headers_with_word if h not in processed_headers]
                if len(unprocessed) > 1:
                    common_word_groups.append(unprocessed)
                    processed_headers.update(unprocessed)
        
        return common_word_groups
    
    def calculate_similarity(self, str1, str2):
        """
        Calculate string similarity using multiple metrics
        and return the highest score
        """
        similarity_scores = []
        
        # Method 1: SequenceMatcher (if enabled)
        if self.use_string_similarity:
            seq_similarity = difflib.SequenceMatcher(None, str1, str2).ratio()
            similarity_scores.append(seq_similarity)
        
        # Method 2: Word overlap coefficient (if enabled)
        if self.use_word_overlap:
            words1 = set(str1.split())
            words2 = set(str2.split())
            
            if words1 and words2:
                intersection = words1.intersection(words2)
                smaller_set = min(len(words1), len(words2))
                word_overlap = len(intersection) / smaller_set
                similarity_scores.append(word_overlap)
        
        # Method 3: Levenshtein-based similarity (if enabled)
        if self.use_string_similarity:
            lev_distance = self.levenshtein_distance(str1, str2)
            max_length = max(len(str1), len(str2))
            if max_length > 0:
                lev_similarity = 1 - (lev_distance / max_length)
                similarity_scores.append(lev_similarity)
        
        # Method 4: Abbreviation detection (if enabled)
        if self.detect_abbreviations:
            abbr_similarity = self.check_abbreviation(str1, str2)
            if abbr_similarity > 0:
                similarity_scores.append(abbr_similarity)
        
        # Return the highest similarity score, or 0 if no methods were used
        return max(similarity_scores) if similarity_scores else 0
    
    def check_abbreviation(self, str1, str2):
        """Check if one string could be an abbreviation of the other"""
        # If one string is much shorter than the other, it might be an abbreviation
        if len(str1) < len(str2) * 0.5:
            shorter, longer = str1, str2
        elif len(str2) < len(str1) * 0.5:
            shorter, longer = str2, str1
        else:
            return 0  # Similar lengths, not an abbreviation situation
        
        # Check if shorter is made up of first letters of longer
        words = longer.split()
        if len(words) >= 2:
            # Get first letters of each word in longer string
            first_letters = ''.join(word[0] for word in words if word)
            
            # If shorter string is similar to the first letters
            if shorter == first_letters:
                return 0.95  # Very high similarity
            elif shorter in first_letters or first_letters in shorter:
                return 0.8   # High similarity
            else:
                lev_similarity = 1 - (self.levenshtein_distance(shorter, first_letters) / max(len(shorter), len(first_letters)))
                if lev_similarity >= 0.7:
                    return lev_similarity
        
        return 0  # Not an abbreviation
    
    def levenshtein_distance(self, s1, s2):
        """Calculate the Levenshtein distance between two strings"""
        if len(s1) < len(s2):
            return self.levenshtein_distance(s2, s1)
        
        # If s2 is empty, the distance is the length of s1
        if len(s2) == 0:
            return len(s1)
        
        previous_row = range(len(s2) + 1)
        for i, c1 in enumerate(s1):
            current_row = [i + 1]
            for j, c2 in enumerate(s2):
                insertions = previous_row[j + 1] + 1
                deletions = current_row[j] + 1
                substitutions = previous_row[j] + (c1 != c2)
                current_row.append(min(insertions, deletions, substitutions))
            previous_row = current_row
        
        return previous_row[-1]
    
    def analyze_and_suggest_merges(self, headers):
        """
        Comprehensive analysis of headers, suggesting potential merges
        based on multiple detection strategies
        """
        results = {
            'exact_duplicates': self.find_exact_duplicates(headers),
            'similar_groups': self.find_similar_groups(headers),
            'common_word_groups': self.find_common_word_headers(headers)
        }
        
        # Generate human-readable suggestions
        suggestions = []
        
        # Exact duplicates have highest priority
        if results['exact_duplicates']:
            suggestions.append("EXACT DUPLICATES (highest priority to merge):")
            for norm, dupes in results['exact_duplicates'].items():
                suggestions.append(f"  • {', '.join(dupes)} → suggest merging to '{dupes[0]}'")
        
        # Similar headers by string similarity
        if results['similar_groups']:
            suggestions.append("\nSIMILAR HEADERS (possible spelling variations):")
            for i, group in enumerate(results['similar_groups']):
                suggestions.append(f"  • Group {i+1}: {', '.join(group)}")
                # Suggest the shortest name or the one that appears first
                suggested_name = min(group, key=len)
                suggestions.append(f"    → suggest merging to '{suggested_name}'")
        
        # Headers with common significant words
        if results['common_word_groups']:
            suggestions.append("\nHEADERS SHARING COMMON WORDS:")
            for i, group in enumerate(results['common_word_groups']):
                suggestions.append(f"  • Group {i+1}: {', '.join(group)}")
                # Extract common words to suggest a name
                common_words = self.extract_common_words(group)
                if common_words:
                    suggestions.append(f"    → suggest merging to '{' '.join(common_words)}'")
        
        return results, '\n'.join(suggestions)
    
    def extract_common_words(self, headers):
        """Extract common words from a group of headers"""
        if not headers:
            return []
            
        # Get all words from the first header
        norm_headers = [self.normalize_text(h) for h in headers]
        words_sets = [set(h.split()) for h in norm_headers if h]
        
        if not words_sets:
            return []
            
        # Find intersection of all word sets
        common_words = set.intersection(*words_sets) if words_sets else set()
        
        # Sort words by their position in the first header
        if common_words and norm_headers:
            first_header_words = norm_headers[0].split()
            return sorted(common_words, key=lambda w: first_header_words.index(w) if w in first_header_words else 999)
        
        return []