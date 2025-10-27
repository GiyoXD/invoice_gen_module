# Example: Processor with CUSTOM sequence (not tied to LayoutBuilder)

class CustomSequenceProcessor(SheetProcessor):
    """
    Example processor that defines its OWN sequence,
    not tied to LayoutBuilder's pre-defined order.
    """
    
    def process(self) -> bool:
        # OPTION 1: Call builders directly in ANY order you want
        
        # 1. First capture template state (your choice)
        template_state = TemplateStateBuilder(
            worksheet=self.worksheet,
            num_header_cols=len(self.sheet_config['header_to_write']),
            header_end_row=18,
            footer_start_row=21
        )
        
        # 2. Maybe do text replacement AFTER template capture (different from LayoutBuilder!)
        text_replacer = TextReplacementBuilder(
            workbook=self.workbook,
            invoice_data=self.invoice_data
        )
        text_replacer.build()
        
        # 3. Write data BEFORE header (completely different sequence!)
        data_builder = DataTableBuilder(...)
        data_builder.build()
        
        # 4. Then write header
        header_builder = HeaderBuilder(...)
        header_info = header_builder.build()
        
        # 5. Then restore template
        template_state.restore_state(...)
        
        # 6. Then write footer separately
        footer_builder = FooterBuilder(...)
        footer_builder.build()
        
        return True
    
    # OPTION 2: Use LayoutBuilder but override its behavior
    def process_with_override(self) -> bool:
        # Use LayoutBuilder but then do custom stuff after
        layout_builder = LayoutBuilder(...)
        success = layout_builder.build()
        
        # Now add YOUR custom sequence AFTER LayoutBuilder
        if success:
            # Add summary section
            summary_builder = SummaryBuilder(...)  # Your custom builder
            summary_builder.build()
            
            # Add signature section
            signature_builder = SignatureBuilder(...)  # Your custom builder
            signature_builder.build()
        
        return success
    
    # OPTION 3: Mix LayoutBuilder with direct builder calls
    def process_hybrid(self) -> bool:
        # Manually capture template state first
        template_state = TemplateStateBuilder(...)
        
        # Use LayoutBuilder for header + data (but not footer)
        layout_builder = LayoutBuilder(
            ...,
            enable_text_replacement=False  # Control what LayoutBuilder does
        )
        success = layout_builder.build()
        
        # Then do custom footer logic
        custom_footer_builder = CustomFooterBuilder(...)
        custom_footer_builder.build()
        
        # Manually restore template
        template_state.restore_state(...)
        
        return success
